import threading
import time
import traceback
from typing import Callable


def do_something_on_thread(
        task: Callable[[], None] | None,
        before_starting: Callable[[], None] | None = None,
        after_task: Callable[[], None] | None = None,
        on_error_log: str = None,
        on_thread: bool = True
):
    """
    Does something on a Thread (e.g. bulk download)
    Args:
        on_thread: Allows you to use this method to execute a non-threaded task
        on_error_log: What to print in case of an exception
        task: The task to do
        before_starting: Something to do before starting the task
        after_task: Something to do after the task is finished or if an exception is thrown
    """
    if before_starting is not None:
        before_starting()

    def thread_task():
        try:
            task()
        except:
            if on_error_log is not None:
                print(on_error_log)
            traceback.print_exc()
        finally:
            if after_task is not None:
                after_task()

    if on_thread:
        t = threading.Thread(target=thread_task)
        t.start()
    else:
        thread_task()


class CancellationToken:
    def __init__(self):
        self.is_cancelled = False

    def cancel(self):
        self.is_cancelled = True


def do_something_periodically_on_thread(
        task: Callable[[], None] | None,
        interval_s: int,
        cancellation_token: CancellationToken | None = None,
        before_starting: Callable[[], None] | None = None,
        after_task: Callable[[], None] | None = None,
        on_error_log: str = None) -> CancellationToken:
    """
    Does something periodically on a Thread (e.g. bulk download). Returns a cancellation token that can be used
    to stop the thread loop
    Args:
        cancellation_token: The cancellation token to (re-)use. If none is supplied, a new one is created
        interval_s: Seconds interval to repeat the action
        on_error_log: What to print in case of an exception
        task: The task to do
        before_starting: Something to do before starting the task
        after_task: Something to do after the task is finished or if an exception is thrown
    """
    if cancellation_token is None:
        cancellation_token = CancellationToken()

    def thread_task():
        while not cancellation_token.is_cancelled:
            do_something_on_thread(
                task=task,
                before_starting=before_starting,
                after_task=after_task,
                on_error_log=on_error_log,
                on_thread=False
            )
            time.sleep(interval_s)

    print(f'Starting thread with task loop {task} ({interval_s}s)')
    t = threading.Thread(target=thread_task)
    t.start()
    return cancellation_token
