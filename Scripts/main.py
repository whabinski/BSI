import sys
from PyQt5.QtWidgets import QApplication
from ui import MyApp  

def main():
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
    
    
#pyinstaller --onefile --windowed --add-data "Assets/*;Assets" --add-data "tally.txt;." Scripts/main.py
