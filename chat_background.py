"""One bounded optional memory update per user; never holds the chat admission gate."""
import asyncio
_tasks = {}

def schedule_memory(uid, update):
    task = _tasks.get(uid)
    if task is not None and not task.done():
        return
    if len(_tasks) >= 8:
        return
    async def run():
        try:
            await asyncio.wait_for(update(uid), timeout=30)
        except Exception as error:
            print('chat_memory reason=' + type(error).__name__, flush=True)
        finally:
            if _tasks.get(uid) is asyncio.current_task():
                _tasks.pop(uid, None)
    _tasks[uid] = asyncio.create_task(run())
