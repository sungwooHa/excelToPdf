@echo off

:: 가상 환경 생성 및 활성화
if not exist venv (
    echo 가상 환경 생성 중...
    python -m venv venv
)

echo 가상 환경 활성화 중...
call venv\Scripts\activate

:: 필요한 패키지 설치
echo 필요한 패키지 설치 중...
pip install --upgrade setuptools wheel
pip install --upgrade pyinstaller
pip install --upgrade pywin32
pip install --upgrade packaging

:: PyInstaller로 실행 파일 생성
echo 실행 파일 생성 중...
pyinstaller --clean ^
            --onefile ^
            --windowed ^
            --hidden-import=pkg_resources.py2_warn ^
            --hidden-import=pkg_resources.extern ^
            --hidden-import=setuptools ^
            --hidden-import=win32com ^
            --hidden-import=win32com.client ^
            --collect-all setuptools ^
            --collect-all win32com ^
            --collect-all pywintypes ^
            --collect-all packaging ^
            excel_to_pdf_gui.py

:: 가상 환경 비활성화
call deactivate

:: 실행 파일이 성공적으로 생성되었는지 확인
if exist "dist\excel_to_pdf_gui.exe" (
    echo 실행 파일이 성공적으로 생성되었습니다: dist\excel_to_pdf_gui.exe
) else (
    echo 실행 파일 생성 실패.
)