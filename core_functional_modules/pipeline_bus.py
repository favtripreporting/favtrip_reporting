# pipeline_bus.py
import queue

_PIPELINE_QUEUE = None

def get_pipeline_queue():
    global _PIPELINE_QUEUE
    if _PIPELINE_QUEUE is None:
        _PIPELINE_QUEUE = queue.Queue()
    return _PIPELINE_QUEUE