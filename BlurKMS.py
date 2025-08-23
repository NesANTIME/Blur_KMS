import os
import sys
import subprocess

def main():
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    python_path = os.path.join(base_path, "resources", "core", "interprete", "python-3.13.7-amd64", "python.exe")
    script_path = os.path.join(base_path, "Blur.py")
    subprocess.run([python_path, script_path], cwd=base_path)

if __name__ == "__main__":
    main()
