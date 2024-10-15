
import datetime
import subprocess



def git_commit_and_push():
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    try:
        subprocess.run(['git', 'add', '.'], check=True)
        subprocess.run(['git', 'commit', '-m', today], check=True)
        subprocess.run(['git', 'push'], check=True)
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while executing git commands: {e}")

git_commit_and_push()