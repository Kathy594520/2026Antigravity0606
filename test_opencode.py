import subprocess
import sys
import os

def main():
    cmd_path = r"C:\Users\USER\AppData\Local\nodejs\node-v22.14.0-win-x64\opencode.cmd"
    if not os.path.exists(cmd_path):
        print(f"Error: {cmd_path} does not exist.")
        return
        
    print(f"Starting process: {cmd_path} ...")
    sys.stdout.flush()
    
    try:
        # We run opencode directly without shell=True to allow proper process management
        p = subprocess.Popen(
            [cmd_path, "run", "hello", "--dangerously-skip-permissions"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, # redirect stderr to stdout so we capture everything in one stream
            text=True
        )
        
        # Read output line-by-line as it arrives
        while True:
            line = p.stdout.readline()
            if not line:
                break
            print(f"OUTPUT: {line.strip()}")
            sys.stdout.flush()
            
        p.wait()
        print(f"Process exited with code: {p.returncode}")
    except Exception as e:
        print(f"Exception encountered: {e}")
        sys.stdout.flush()

if __name__ == "__main__":
    main()
