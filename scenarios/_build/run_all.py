"""Dry-run every scenario of a flow with the installed agent and say which ones fail.
usage (from the repo root): python scenarios\\_build\\run_all.py flows\\ΑΠΔ.py scenarios\\ΑΠΔ"""
import subprocess, glob, os, sys
flow, folder = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..")
ok, bad = 0, []
for sc in sorted(glob.glob(os.path.join(folder, "*.json"))):
    p = subprocess.run([r"C:\Myrmex\myrmex.exe", flow, "--dry-run", "--scenario", sc], capture_output=True,
                       text=True, encoding="utf-8", errors="replace", cwd=root)
    out = p.stdout + p.stderr
    if p.returncode == 0 and "MISSED" not in out and "[FAIL]" not in out:
        ok += 1
    else:
        got = [l.strip() for l in out.splitlines() if l.strip().startswith("→")][-1:]
        why = [l.strip() for l in out.splitlines() if "[FAIL]" in l][:1]
        bad.append((os.path.basename(sc), got, why))
print(f"{ok}/{ok + len(bad)} OK")
for b in bad:
    print("✗", b[0], "|", (b[2] or [""])[0][:230], "|", (b[1] or [""])[0][:200])
