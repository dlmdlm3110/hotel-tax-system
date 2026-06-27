#!/usr/bin/env python3
"""デスクトップにショートカットを作成（OneDrive対応・Windows専用）"""
import os, sys, subprocess

if sys.platform != 'win32':
    sys.exit(0)

app_dir  = os.path.dirname(os.path.abspath(__file__))
bat_path = os.path.join(app_dir, 'launch.bat')
ico_path = os.path.join(app_dir, 'appicon.ico')
cmd_exe  = os.path.join(os.environ.get('SystemRoot', r'C:\Windows'), 'System32', 'cmd.exe')

ps_lines = [
    '$sh = New-Object -COM WScript.Shell',
    "$dt = [Environment]::GetFolderPath('Desktop')",          # OneDriveでも正しいパスを返す
    "$lnk = [IO.Path]::Combine($dt, '宿泊税計算システム.lnk')",
    '$s = $sh.CreateShortcut($lnk)',
    f"$s.TargetPath = '{cmd_exe}'",
    f'$s.Arguments = \'/c "{bat_path}"\'',
    f"$s.WorkingDirectory = '{app_dir}'",
    '$s.WindowStyle = 1',
    "$s.Description = '宿泊税計算システム'",
]
if os.path.exists(ico_path):
    ps_lines.append(f"$s.IconLocation = '{ico_path},0'")
ps_lines += [
    '$s.Save()',
    "Write-Host ('[OK] 作成: ' + $lnk)",
]

result = subprocess.run(
    ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
     '\n'.join(ps_lines)],
    capture_output=True, text=True, encoding='utf-8', errors='replace',
)
if result.stdout.strip():
    print(result.stdout.strip())
if result.returncode != 0:
    print('[エラー]', result.stderr.strip(), file=sys.stderr)
    sys.exit(1)
