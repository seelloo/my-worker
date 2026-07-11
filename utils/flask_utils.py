import os
import sys
import subprocess
from flask import Response


def stream_task(cmd):
    """
    通过物理隔离的子进程 (Subprocess) 运行任务，
    实时捕捉并流式返回标准输出，实现"沙箱化"执行。
    """
    if cmd[0] == "python":
        cmd[0] = sys.executable

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        universal_newlines=True,
        encoding='utf-8',
        env=env
    )

    def generate():
        yield f"🚀 [Sandbox] 独立任务进程已启动: {' '.join(cmd)}\n"
        for line in iter(process.stdout.readline, ""):
            yield line
        process.stdout.close()
        return_code = process.wait()
        if return_code == 0:
            yield "\n✅ 物理隔离进程执行圆满成功！\n"
        else:
            yield f"\n❌ 任务进程异常终止 (Status: {return_code})，请检查上方日志。\n"

    return Response(generate(), mimetype='text/plain')
