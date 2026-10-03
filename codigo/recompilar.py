"""Genera el ejecutable y el ZIP de entrega desde el código actual."""
import hashlib
import json
import os
import shutil
import ssl
import struct
import subprocess
import sys
import uuid
import zipfile
from pathlib import Path


BASE = Path(__file__).resolve().parent
SOURCES = ("app.py", "core.py", "registro.py", "Iniciar.cmd", "Compilar.cmd", "recompilar.py",
           "test_core.py", "test_registro.py", "test_app.py", "LEEME.md")


def run(command, env, **kwargs):
    subprocess.run([str(arg) for arg in command], cwd=str(BASE), env=env, check=True, **kwargs)


def main():
    if sys.platform != "win32" or struct.calcsize("P") != 8:
        raise RuntimeError("La compilación requiere Windows y Python de 64 bits.")
    environment = os.environ.copy()
    environment["PYINSTALLER_CONFIG_DIR"] = str(BASE / ".build-env" / "pyinstaller-cache")
    python = BASE / ".build-env" / "Scripts" / "python.exe"
    if not python.exists():
        print("Preparando el entorno de compilación…", flush=True)
        run([sys.executable, "-m", "venv", BASE / ".build-env"], environment)
    available = subprocess.run([str(python), "-c", "import PyInstaller"], cwd=str(BASE),
                               env=environment, capture_output=True)
    if available.returncode:
        print("Instalando PyInstaller. Este paso requiere internet una sola vez…", flush=True)
        certificates = BASE / ".build-env" / "windows-ca.pem"
        certificates.write_text("".join(ssl.DER_cert_to_PEM_cert(cert)
                                        for cert in ssl.create_default_context().get_ca_certs(binary_form=True)),
                                encoding="ascii")
        run([python, "-m", "pip", "--isolated", "install", "--index-url", "https://pypi.org/simple",
             "--cert", certificates, "--disable-pip-version-check", "pyinstaller==6.16.0"], environment)

    source_bytes = {name: (BASE / name).read_bytes() for name in SOURCES}
    print("\n1/4  Verificando cálculos e interfaz…", flush=True)
    run([python, "-m", "unittest", "-v", "test_core", "test_registro", "test_app"], environment)

    print("\n2/4  Compilando el diseño actual…", flush=True)
    staging = BASE / "build" / "release"
    run([python, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile", "--windowed",
         "--name", "eCheck", "--distpath", staging, "app.py"], environment)
    executable = staging / "eCheck.exe"

    print("\n3/4  Comprobando el ejecutable…", flush=True)
    report_path = BASE / "build" / ("validacion-" + uuid.uuid4().hex + ".json")
    run([executable, "--self-test", report_path], environment,
        timeout=60, creationflags=subprocess.CREATE_NO_WINDOW)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    if not report.get("ok") or not report.get("frozen"):
        raise RuntimeError("El ejecutable no pasó la verificación: {}".format(report))
    changed = [name for name, data in source_bytes.items() if (BASE / name).read_bytes() != data]
    if changed:
        raise RuntimeError("Se modificaron archivos durante la compilación: {}. Volvé a compilar.".format(
            ", ".join(changed)))

    print("\n4/4  Generando el ZIP para distribuir…", flush=True)
    delivery = BASE / "entrega"
    delivery.mkdir(exist_ok=True)
    archive = delivery / "eCheck-Windows-x64.zip"
    pending_zip = delivery / "eCheck-Windows-x64.zip.tmp"
    with zipfile.ZipFile(pending_zip, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as handle:
        handle.write(executable, "eCheck/eCheck.exe")
        handle.writestr("eCheck/LEEME.md", source_bytes["LEEME.md"])
        for name, data in source_bytes.items():
            handle.writestr("eCheck/codigo/" + name, data)
    digest = hashlib.sha256(executable.read_bytes()).digest()
    with zipfile.ZipFile(pending_zip) as handle:
        if handle.testzip() is not None or hashlib.sha256(handle.read("eCheck/eCheck.exe")).digest() != digest:
            raise RuntimeError("No se pudo verificar la integridad del ZIP.")
    dist = BASE / "dist"
    dist.mkdir(exist_ok=True)
    pending_exe = dist / "eCheck.exe.tmp"
    shutil.copyfile(executable, pending_exe)
    pending_exe.replace(dist / "eCheck.exe")
    pending_zip.replace(archive)
    print("\nLISTO. Ejecutable y ZIP actualizados:", flush=True)
    print(dist / "eCheck.exe", flush=True)
    print(archive, flush=True)
    print("\nExtraé el ZIP en las otras PC y abrí el nuevo eCheck.exe.", flush=True)


if __name__ == "__main__":
    try:
        main()
    except subprocess.TimeoutExpired:
        print("\nERROR: La prueba del ejecutable superó 60 segundos. No se reemplazó el ZIP anterior.", file=sys.stderr)
        sys.exit(1)
    except (OSError, RuntimeError, subprocess.CalledProcessError, ValueError, zipfile.BadZipFile) as exc:
        print("\nERROR: {}\nNo se reemplazó el ZIP anterior. Si eCheck está abierto, cerralo y volvé a intentar.".format(exc), file=sys.stderr)
        sys.exit(1)
