pyinstaller --clean --onefile --windowed ^
    --distpath="app/dist/" ^
    --icon="icons/icon.ico" ^
    --name="Labor Utility v0.2" ^
    app/app/app.py