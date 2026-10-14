"""
solver_io.py — launch the REMFluor Fortran solver WITHOUT cmd.exe.

v113: Previously every caller did

    subprocess.Popen('"exe" < "input.inp" > "output.out"',
                     shell=True, cwd=model_dir)

shell=True routes through cmd.exe, and cmd.exe cannot use a UNC path
(\\\\server\\share\\...) as its current directory — it silently falls back
to C:\\Windows.  The solver then started in the wrong folder, never found
input.inp, and the run failed whenever the model lived on the network
server.  (Save / Load / Visualize kept working because they are pure
Python and use absolute paths.)

Here Python opens input.inp / output.out itself and hands the handles to
the solver directly (CreateProcess), which DOES accept UNC working
directories.  Output files are byte-identical to the old shell version.

Side benefit: a timeout now kills the solver itself — with shell=True it
only killed cmd.exe and left the Fortran process running.
"""
from __future__ import annotations
import os
import subprocess

INPUT_NAME = "input.inp"
OUTPUT_NAME = "output.out"


def _no_window_kwargs() -> dict:
    """Keep the console solver hidden.  shell=True used to hide cmd.exe
    (SW_HIDE) and the solver inherited that hidden console; without the
    shell, a console app spawned from the windowed .exe would otherwise
    pop up its own black window."""
    if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
        return {"creationflags": subprocess.CREATE_NO_WINDOW}
    return {}


def start_solver(exe_path: str, work_dir: str,
                 inp_name: str = INPUT_NAME,
                 out_name: str = OUTPUT_NAME,
                 stderr=subprocess.DEVNULL) -> subprocess.Popen:
    """Start  exe < work_dir/inp_name > work_dir/out_name  in work_dir.

    Non-blocking: returns the Popen so the caller can poll() it.
    Equivalent of the old  '"exe" < "input.inp" > "output.out"'  shell line.
    """
    inp_path = os.path.join(work_dir, inp_name)
    out_path = os.path.join(work_dir, out_name)
    with open(inp_path, "rb") as fin, open(out_path, "wb") as fout:
        # The child gets its own duplicated handles; closing ours right
        # after spawn (end of the with-block) is safe.
        return subprocess.Popen([exe_path], stdin=fin, stdout=fout,
                                stderr=stderr, cwd=work_dir,
                                **_no_window_kwargs())


def run_solver(exe_path: str, work_dir: str, timeout=None,
               inp_name: str = INPUT_NAME,
               out_name: str = OUTPUT_NAME) -> subprocess.CompletedProcess:
    """Blocking version (calibration / sensitivity loops).

    Mirrors the old  subprocess.run(cmd, shell=True, cwd=..., timeout=...,
    capture_output=True)  call: stdout goes to output.out, stderr is
    captured.  Raises subprocess.TimeoutExpired after killing the solver.
    """
    proc = start_solver(exe_path, work_dir, inp_name, out_name,
                        stderr=subprocess.PIPE)
    try:
        _, err = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        try:
            proc.communicate(timeout=10)
        except subprocess.TimeoutExpired:
            pass
        raise
    return subprocess.CompletedProcess([exe_path], proc.returncode,
                                       stdout=b"", stderr=err or b"")
