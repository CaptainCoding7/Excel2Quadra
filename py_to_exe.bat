python -m PyInstaller --noconfirm --log-level=WARN ^
    --onefile --noconsole ^
    --hidden-import=pandas ^
    --hidden-import=datetime ^
    --hidden-import=enum ^
    --hidden-import=os ^
    main.py