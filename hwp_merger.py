import sys
import os
import zipfile
import shutil
import tempfile
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem, QFileDialog,
    QLabel, QProgressBar, QMessageBox, QAbstractItemView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QDragEnterEvent, QDropEvent


class MergeWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, file_paths: list, output_path: str):
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path

    def run(self):
        try:
            # 확장자 확인
            extensions = {Path(f).suffix.lower() for f in self.file_paths}
            
            if extensions == {'.hwpx'} or (extensions <= {'.hwpx'}):
                self.merge_hwpx_files()
            elif '.hwp' in extensions:
                self.merge_with_hwp_com()
            else:
                self.merge_hwpx_files()
        except Exception as e:
            self.finished.emit(False, str(e))

    def merge_hwpx_files(self):
        """HWPX 파일 병합 (ZIP 기반 XML 처리)"""
        import xml.etree.ElementTree as ET
        
        self.progress.emit(10, "첫 번째 파일 읽는 중...")
        
        base_file = self.file_paths[0]
        temp_dir = tempfile.mkdtemp()
        
        try:
            base_dir = os.path.join(temp_dir, "base")
            os.makedirs(base_dir)
            
            with zipfile.ZipFile(base_file, 'r') as z:
                z.extractall(base_dir)
            
            # 기준 파일의 content.hpf (본문) 파싱
            content_xml_path = os.path.join(base_dir, "Contents", "content.xml")
            if not os.path.exists(content_xml_path):
                # 다른 경로 시도
                for root, dirs, files in os.walk(base_dir):
                    for f in files:
                        if f == "content.xml":
                            content_xml_path = os.path.join(root, f)
                            break
            
            ET.register_namespace('', 'urn:schemas-microsoft-com:office:spreadsheet')
            
            base_tree = ET.parse(content_xml_path)
            base_root = base_tree.getroot()
            
            # 네임스페이스 추출
            ns = self._get_namespaces(content_xml_path)
            
            # body 섹션 찾기
            body_elem = self._find_body(base_root, ns)
            
            total = len(self.file_paths)
            
            for i, filepath in enumerate(self.file_paths[1:], 1):
                self.progress.emit(
                    int(10 + (i / total) * 80),
                    f"파일 병합 중: {Path(filepath).name}"
                )
                
                merge_dir = os.path.join(temp_dir, f"merge_{i}")
                os.makedirs(merge_dir)
                
                with zipfile.ZipFile(filepath, 'r') as z:
                    z.extractall(merge_dir)
                
                merge_xml = os.path.join(merge_dir, "Contents", "content.xml")
                if not os.path.exists(merge_xml):
                    for root_d, dirs, files in os.walk(merge_dir):
                        for f in files:
                            if f == "content.xml":
                                merge_xml = os.path.join(root_d, f)
                                break
                
                merge_tree = ET.parse(merge_xml)
                merge_root = merge_tree.getroot()
                merge_body = self._find_body(merge_root, ns)
                
                # 섹션 추가 (페이지 나누기 포함)
                if body_elem is not None and merge_body is not None:
                    # 섹션 구분자 추가
                    for child in list(merge_body):
                        body_elem.append(child)
                
                # 이미지 등 리소스 복사
                self._copy_resources(merge_dir, base_dir)
            
            self.progress.emit(90, "출력 파일 생성 중...")
            
            # 수정된 XML 저장
            base_tree.write(content_xml_path, encoding='utf-8', xml_declaration=True)
            
            # 새 HWPX로 압축
            output_path = self.output_path
            if not output_path.lower().endswith('.hwpx'):
                output_path += '.hwpx'
            
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for root_d, dirs, files in os.walk(base_dir):
                    for file in files:
                        file_path = os.path.join(root_d, file)
                        arcname = os.path.relpath(file_path, base_dir)
                        zout.write(file_path, arcname)
            
            self.progress.emit(100, "완료!")
            self.finished.emit(True, output_path)
            
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def merge_with_hwp_com(self):
        """한글 COM 자동화를 이용한 HWP 병합 (한글 설치 필요)"""
        try:
            import win32com.client
        except ImportError:
            self.finished.emit(False, "win32com이 없습니다. pywin32를 설치하거나 HWPX 파일을 사용하세요.")
            return
        
        self.progress.emit(10, "한글 프로그램 시작 중...")
        
        hwp = None
        try:
            hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            hwp.XHwpWindows.Item(0).Visible = False
            
            # 첫 번째 파일 열기
            first_file = os.path.abspath(self.file_paths[0])
            hwp.Open(first_file)
            
            total = len(self.file_paths)
            
            for i, filepath in enumerate(self.file_paths[1:], 1):
                self.progress.emit(
                    int(10 + (i / total) * 80),
                    f"파일 삽입 중: {Path(filepath).name}"
                )
                
                # 문서 끝으로 이동
                hwp.Run("MoveDocEnd")
                
                # 페이지 나누기 삽입
                hwp.Run("BreakPage")
                
                # 파일 삽입
                abs_path = os.path.abspath(filepath)
                act = hwp.CreateAction("InsertFile")
                set = act.CreateSet()
                set.SetItem("FileName", abs_path)
                set.SetItem("KeepSection", 1)
                set.SetItem("KeepCharshape", 1)
                set.SetItem("KeepParashape", 1)
                set.SetItem("KeepStyle", 1)
                act.Execute(set)
            
            self.progress.emit(90, "저장 중...")
            
            output_path = self.output_path
            if not output_path.lower().endswith(('.hwp', '.hwpx')):
                output_path += '.hwp'
            
            hwp.SaveAs(os.path.abspath(output_path), "HWP")
            hwp.Quit()
            
            self.progress.emit(100, "완료!")
            self.finished.emit(True, output_path)
            
        except Exception as e:
            if hwp:
                try:
                    hwp.Quit()
                except:
                    pass
            self.finished.emit(False, f"한글 COM 오류: {str(e)}")

    def _get_namespaces(self, xml_path):
        """XML에서 네임스페이스 추출"""
        ns = {}
        with open(xml_path, 'rb') as f:
            content = f.read().decode('utf-8', errors='ignore')
            import re
            for match in re.finditer(r'xmlns:?(\w*)=["\']([^"\']+)["\']', content):
                prefix, uri = match.group(1), match.group(2)
                ns[prefix if prefix else 'default'] = uri
        return ns

    def _find_body(self, root, ns):
        """XML에서 body 요소 찾기"""
        # 한글 HWPX 구조에서 본문 섹션 탐색
        for elem in root.iter():
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if tag.lower() in ('body', 'hpf:body', 'sec', 'section'):
                return elem
        # 루트의 첫 번째 주요 자식 반환
        if len(root) > 0:
            return root
        return root

    def _copy_resources(self, src_dir, dst_dir):
        """이미지 등 리소스 파일 복사"""
        resource_dirs = ['BinData', 'Contents', 'Preview']
        for res_dir in resource_dirs:
            src = os.path.join(src_dir, res_dir)
            dst = os.path.join(dst_dir, res_dir)
            if os.path.exists(src):
                if not os.path.exists(dst):
                    os.makedirs(dst)
                for f in os.listdir(src):
                    src_f = os.path.join(src, f)
                    dst_f = os.path.join(dst, f)
                    if os.path.isfile(src_f) and not os.path.exists(dst_f):
                        shutil.copy2(src_f, dst_f)


class DropListWidget(QListWidget):
    """드래그 앤 드롭 지원 리스트"""
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            files = []
            for url in urls:
                path = url.toLocalFile()
                if path.lower().endswith(('.hwp', '.hwpx')):
                    files.append(path)
            if files:
                self.files_dropped.emit(files)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HWP/HWPX 파일 합치기")
        self.setMinimumSize(600, 500)
        self.worker = None
        self._init_ui()

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        # 제목
        title = QLabel("HWP / HWPX 파일 합치기")
        title.setFont(QFont("맑은 고딕", 14, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # 안내
        hint = QLabel("파일을 추가하거나 여기로 드래그 & 드롭하세요. 위아래로 순서를 바꿀 수 있습니다.")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(hint)

        # 파일 리스트
        self.list_widget = DropListWidget()
        self.list_widget.setDragDropMode(QAbstractItemView.DragDrop)
        self.list_widget.setDefaultDropAction(Qt.MoveAction)
        self.list_widget.files_dropped.connect(self.add_files_to_list)
        layout.addWidget(self.list_widget)

        # 버튼 행
        btn_layout = QHBoxLayout()
        
        self.btn_add = QPushButton("파일 추가")
        self.btn_add.clicked.connect(self.add_files)
        
        self.btn_remove = QPushButton("선택 삭제")
        self.btn_remove.clicked.connect(self.remove_selected)
        
        self.btn_clear = QPushButton("전체 삭제")
        self.btn_clear.clicked.connect(self.list_widget.clear)
        
        self.btn_up = QPushButton("▲ 위로")
        self.btn_up.clicked.connect(self.move_up)
        
        self.btn_down = QPushButton("▼ 아래로")
        self.btn_down.clicked.connect(self.move_down)
        
        for btn in [self.btn_add, self.btn_remove, self.btn_clear, self.btn_up, self.btn_down]:
            btn_layout.addWidget(btn)
        
        layout.addLayout(btn_layout)

        # 파일 수 표시
        self.lbl_count = QLabel("파일 0개 선택됨")
        self.lbl_count.setAlignment(Qt.AlignRight)
        self.list_widget.model().rowsInserted.connect(self._update_count)
        self.list_widget.model().rowsRemoved.connect(self._update_count)
        layout.addWidget(self.lbl_count)

        # 진행바
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_status)

        # 병합 버튼
        self.btn_merge = QPushButton("파일 합치기")
        self.btn_merge.setFont(QFont("맑은 고딕", 12, QFont.Bold))
        self.btn_merge.setMinimumHeight(45)
        self.btn_merge.setStyleSheet(
            "QPushButton { background-color: #0078d4; color: white; border-radius: 6px; }"
            "QPushButton:hover { background-color: #106ebe; }"
            "QPushButton:disabled { background-color: #cccccc; }"
        )
        self.btn_merge.clicked.connect(self.start_merge)
        layout.addWidget(self.btn_merge)

    def _update_count(self):
        count = self.list_widget.count()
        self.lbl_count.setText(f"파일 {count}개")

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "HWP/HWPX 파일 선택", "",
            "한글 파일 (*.hwp *.hwpx);;모든 파일 (*)"
        )
        self.add_files_to_list(files)

    def add_files_to_list(self, files):
        existing = [self.list_widget.item(i).data(Qt.UserRole)
                    for i in range(self.list_widget.count())]
        for f in files:
            if f not in existing:
                item = QListWidgetItem(Path(f).name)
                item.setData(Qt.UserRole, f)
                item.setToolTip(f)
                self.list_widget.addItem(item)

    def remove_selected(self):
        for item in self.list_widget.selectedItems():
            self.list_widget.takeItem(self.list_widget.row(item))

    def move_up(self):
        row = self.list_widget.currentRow()
        if row > 0:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row - 1, item)
            self.list_widget.setCurrentRow(row - 1)

    def move_down(self):
        row = self.list_widget.currentRow()
        if row < self.list_widget.count() - 1:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row + 1, item)
            self.list_widget.setCurrentRow(row + 1)

    def start_merge(self):
        count = self.list_widget.count()
        if count < 2:
            QMessageBox.warning(self, "경고", "파일을 2개 이상 추가해주세요.")
            return

        file_paths = [self.list_widget.item(i).data(Qt.UserRole)
                      for i in range(count)]

        # 저장 경로
        ext = ".hwpx" if all(f.lower().endswith('.hwpx') for f in file_paths) else ".hwp"
        output_path, _ = QFileDialog.getSaveFileName(
            self, "저장할 파일 선택", f"merged{ext}",
            "한글 파일 (*.hwp *.hwpx)"
        )
        if not output_path:
            return

        self.btn_merge.setEnabled(False)
        self.btn_add.setEnabled(False)
        self.progress_bar.setValue(0)

        self.worker = MergeWorker(file_paths, output_path)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def on_progress(self, value, msg):
        self.progress_bar.setValue(value)
        self.lbl_status.setText(msg)

    def on_finished(self, success, message):
        self.btn_merge.setEnabled(True)
        self.btn_add.setEnabled(True)
        if success:
            self.lbl_status.setText("병합 완료!")
            QMessageBox.information(self, "완료", f"파일이 저장되었습니다:\n{message}")
        else:
            self.lbl_status.setText("오류 발생")
            QMessageBox.critical(self, "오류", f"병합 중 오류가 발생했습니다:\n{message}")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
