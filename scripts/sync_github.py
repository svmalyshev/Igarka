"""Commit completed local changes and synchronize Igarka with GitHub.

No force-push, rebasing or automatic conflict resolution. Exit unsuccessfully
when remote edits and local edits require a human decision.
"""
import fcntl
import os
from pathlib import Path
import subprocess
from datetime import datetime, timezone

ROOT = Path(__file__).resolve().parents[1]
EXPECTED_REMOTE = 'https://github.com/svmalyshev/Igarka.git'


def git(*args, capture=True):
    environment = os.environ.copy()
    environment['GIT_TERMINAL_PROMPT'] = '0'
    result = subprocess.run(['git', *args], cwd=ROOT, env=environment,
                            text=True, capture_output=capture)
    if result.returncode:
        raise RuntimeError((result.stderr or result.stdout or 'Git failed').strip())
    return result.stdout.strip() if capture else ''


def main():
    assert git('rev-parse', '--show-toplevel') == str(ROOT)
    assert git('branch', '--show-current') == 'main', 'Sync requires branch main'
    assert git('remote', 'get-url', 'origin') == EXPECTED_REMOTE
    gitdir=ROOT/'.git'
    with (gitdir/'igarka-sync.lock').open('a') as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            print('Another sync is already running.')
            return
        for name in ('MERGE_HEAD','CHERRY_PICK_HEAD','REBASE_HEAD','rebase-merge','rebase-apply','index.lock'):
            if (gitdir/name).exists():
                raise RuntimeError('Git operation in progress: '+name)
        git('fetch','origin','main')
        ahead,behind=map(int,git('rev-list','--left-right','--count','HEAD...origin/main').split())
        dirty=bool(git('status','--porcelain'))
        if behind:
            if ahead or dirty:
                raise RuntimeError('GitHub has new commits and local changes exist. Review both versions before sync.')
            git('merge','--ff-only','origin/main')
        # A completed HTML build must exist; raw datasets stay outside this repo.
        if not (ROOT/'index.html').is_file():
            raise RuntimeError('Missing index.html')
        candidates=git('ls-files','--cached','--others','--exclude-standard','-z').split('\0')
        for relative in filter(None,candidates):
            file=ROOT/relative
            if file.is_file() and file.stat().st_size>=100*1024*1024:
                raise RuntimeError('File exceeds regular GitHub limit: '+relative)
        if git('status','--porcelain'):
            git('add','--all')
            staged=git('diff','--cached','--name-only')
            if staged:
                git('commit','-m','Update Igarka portal '+datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC'))
        git('push','origin','main')
        print('Igarka is synchronized with GitHub.')


if __name__=='__main__':
    try:
        main()
    except (RuntimeError,AssertionError) as error:
        raise SystemExit(str(error))
