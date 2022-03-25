# AutoPPTMaker
Automated worship slide creator for FCNABC's English Congregation.

Refer to "Auto PPT Maker Instructions.docx" for more info.

-------------------

Project Setup:
1. Install poetry via "pip install --upgrade poetry"
2. Navigate to project directory (where this README file is)
3. Install virtual environment via "poetry update"

To run projects, enter the virtual environment via "poetry shell" (avoids the usage of "poetry run") or execute directly via "poetry run python [python file]".
The main application can be executed via "poetry run python .\AutoPPTMaker\SlideMaker.py".

-------------------

To export as .exe "poetry run pyinstaller --onefile .\AutoPPTMaker\SlideMaker.py --path=".\AutoPPTMaker"".