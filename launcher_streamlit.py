import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()
