from __future__ import annotations
import subprocess
import sys

cmd=[sys.executable,'-m','pytest','-q','tests/test_attendance_optional_floater.py','tests/test_attendance_targets.py']
raise SystemExit(subprocess.call(cmd))
