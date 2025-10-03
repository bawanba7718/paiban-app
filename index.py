import subprocess
import sys
import os

def install_requirements():
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

if __name__ == "__main__":
    install_requirements()
    # 启动Streamlit应用
    os.system("streamlit run paiban.py --server.port=$PORT --server.address=0.0.0.0")
