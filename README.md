# AutoPPTMaker

Automated worship slide creator for FCNABC's English Congregation.

Refer to "Auto PPT Maker Instructions.docx" for user guide.

## 1. Project Setup

1. Open a command console

2. Install poetry via

    ```bash
    pip install --upgrade poetry
    ```

3. Navigate to project directory (where this README file is)

4. Install virtual environment via:

    ```bash
    poetry update
    ```

## 2. Executing Project

### 2.1 General

To run a Python file, either:

1. Enter the virtual environment (avoids the usage of "poetry run") via the following command and then execute python as normally (i.e., *python [python file]*)

    ```bash
    poetry shell
    ```

2. Or execute directly via

    ```bash
    poetry run python [python file]
    ```

### 2.2 Main Application

After navigating to the AutoPPTMaker subfolder, the main application can be executed via:

```bash
poetry run python SlideMaker.py
```

## 3. Project Export

To export as .exe, use pyinstaller via:

```bash
poetry run pyinstaller --onefile .\AutoPPTMaker\SlideMaker.py --path=".\AutoPPTMaker"
```
