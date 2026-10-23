"""CI-проверка готового .app после распаковки штатным средством macOS."""

from pathlib import Path
import os
import signal
import subprocess
import time


root = Path("dist/roundtrip").resolve()
root.mkdir()
subprocess.run(["ditto", "-x", "-k", "dist/YMMatcher-M1.zip", str(root)], check=True)
app = root / "YMMatcher.app"
exe = app / "Contents/MacOS/YMMatcher"
assert os.access(exe, os.X_OK), "Потеряны права запуска"
for entry in app.rglob("*"):
    if entry.is_symlink():
        assert entry.exists(), f"Повреждена символическая ссылка: {entry}"
subprocess.run(["codesign", "--verify", "--deep", "--strict", str(app)], check=True)
subprocess.run(["plutil", "-lint", str(app / "Contents/Info.plist")], check=True)


def launch_and_check(command, finder=False):
    report = root / "Ошибка запуска - YM.txt"
    report.unlink(missing_ok=True)
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)
    try:
        deadline = time.monotonic() + 20
        while not report.exists() and process.poll() is None and time.monotonic() < deadline:
            time.sleep(0.2)
        if not report.exists():
            if process.poll() is None:
                os.killpg(process.pid, signal.SIGTERM)
            stdout, stderr = process.communicate(timeout=5)
            raise AssertionError(f"Не достигнут обработчик программы. stdout={stdout!r}; stderr={stderr!r}")
        text = report.read_text(encoding="utf-8")
        assert "FileNotFoundError" in text and "zvonko_in.xlsx" in text, text
        print("PASS:", "LaunchServices (open .app)" if finder else "прямой запуск исполняемого файла",
              "— создано окно Tk и достигнута проверка входных файлов", flush=True)
    finally:
        if finder:
            subprocess.run(["pkill", "-f", str(exe)], check=False)
        if process.poll() is None:
            os.killpg(process.pid, signal.SIGTERM)
        process.communicate(timeout=5)


launch_and_check([str(exe)])
launch_and_check(["open", "-n", "-W", str(app)], finder=True)
print("PASS: архив, права запуска, символические ссылки, подпись и два способа запуска.")
