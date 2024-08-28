编译方法：

nuitka编译指令：
nuitka --mingw64 --standalone --onefile --jobs=12 --lto=yes --show-progress --enable-plugin=tk-inter --remove-output --windows-console-mode=disable main.py

pyinstaller编译指令：
pyinstaller --onefile --noconsole .\main.py
