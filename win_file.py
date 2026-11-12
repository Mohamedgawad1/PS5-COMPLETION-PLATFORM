# -*- coding: utf-8 -*-
"""Lock-safe file helpers for the Windows build chain.

Excel keeps a workbook write-locked while it is open:

    PermissionError: [WinError 32] The process cannot access the file
    because it is being used by another process

The old "restore the backup" path copied over that locked file, raised the
same WinError 32 again, and turned a temporary "please close Excel" into a
failed build. Every helper here either succeeds or leaves the target
byte-for-byte untouched and returns False, so the caller can retry next cycle.
"""
from __future__ import annotations

import errno
import os
import shutil
import time
import zipfile

# WinError 5    = access denied      (OneDrive / antivirus still writing)
# WinError 32   = sharing violation  (Excel has the file open)
# WinError 33   = lock violation     (another process holds a byte lock)
# WinError 1224 = user mapped file   (file mapped in another process)
_LOCK_WINERRORS = {5, 32, 33, 1224}
_LOCK_ERRNOS = {errno.EACCES, errno.EBUSY, errno.EPERM}


def is_lock_error(exc):
    """True for 'file in use / access denied' failures only."""
    if isinstance(exc, PermissionError):
        return True
    winerr = getattr(exc, 'winerror', None)
    if winerr is not None and winerr in _LOCK_WINERRORS:
        return True
    return getattr(exc, 'errno', None) in _LOCK_ERRNOS


def is_locked(path):
    """True when another process (normally Excel) denies write access.

    WinError 32/33/1224 mean "in use". WinError 5 (access denied) is NOT
    treated as a lock here: it usually means a read-only attribute, and
    deferring forever would hide that.
    """
    if not os.path.exists(path):
        return False
    try:
        with open(path, 'r+b'):
            return False
    except OSError as e:
        winerr = getattr(e, 'winerror', None)
        if winerr is not None:
            return winerr in (32, 33, 1224)
        return isinstance(e, PermissionError)


def friendly(path):
    return '"%s"' % os.path.basename(path)


def retry(fn, path='', attempts=15, delay=2.0, max_delay=15.0,
          label=None, log=print):
    """Call fn(), retrying while Windows reports the file as locked.

    Re-raises the last lock error when every attempt is blocked, so callers
    can decide between "defer" and "fail".
    """
    attempts = max(1, attempts)
    last = None
    for i in range(attempts):
        try:
            return fn()
        except OSError as e:
            if not is_lock_error(e):
                raise
            last = e
            if i + 1 < attempts:
                if log:
                    msg = label or ('%s is open in another program - close it'
                                    % friendly(path))
                    log('  [wait] %s (%d/%d)' % (msg, i + 1, attempts))
                time.sleep(min(delay * (1.5 ** i), max_delay))
    raise last


def _remove_quiet(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


def _validate_zip(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        bad = z.testzip()
    if bad is not None:
        raise RuntimeError('corrupt workbook member: %s' % bad)
    if '[Content_Types].xml' not in names or not any(
            n.startswith('xl/') for n in names):
        raise RuntimeError('saved file is not a workbook')


def replace_with_retry(src, dst, attempts=15, delay=2.0, max_delay=15.0,
                       log=print):
    """os.replace() with retries. False (not an exception) when still locked."""
    try:
        retry(lambda: os.replace(src, dst), dst, attempts, delay,
              max_delay, log=log)
        return True
    except OSError as e:
        if is_lock_error(e):
            if log:
                log('  [skip] %s is open in Excel - left untouched, will '
                    'retry next cycle' % friendly(dst))
            return False
        raise


def copy_with_retry(src, dst, attempts=15, delay=2.0, max_delay=15.0,
                    log=print):
    """Atomic copy: write dst.copying.tmp, then swap it into place.

    The destination is never truncated before the copy has fully succeeded,
    so a locked workbook cannot end up half-written.
    """
    tmp = dst + '.copying.tmp'
    try:
        retry(lambda: shutil.copy2(src, tmp), dst, attempts, delay,
              max_delay, log=log)
    except OSError as e:
        _remove_quiet(tmp)
        if is_lock_error(e):
            if log:
                log('  [skip] could not read %s - will retry next cycle'
                    % friendly(src))
            return False
        raise
    if replace_with_retry(tmp, dst, attempts, delay, max_delay, log=log):
        return True
    _remove_quiet(tmp)
    return False


def save_workbook(wb, path, attempts=15, delay=2.0, max_delay=15.0,
                  log=print):
    """Save an openpyxl workbook without ever damaging a locked target.

    tmp file -> zip validation -> atomic replace.

    Returns True on success. On a lock the original file is untouched and
    False is returned; the caller should retry on the next cycle.
    """
    folder = os.path.dirname(os.path.abspath(path))
    tmp = os.path.join(folder, os.path.basename(path) + '.saving.tmp')
    try:
        retry(lambda: wb.save(tmp), tmp, attempts, delay, max_delay, log=log)
    except OSError as e:
        _remove_quiet(tmp)
        if is_lock_error(e):
            if log:
                log('  [skip] %s is open in Excel - left untouched, will '
                    'retry next cycle' % friendly(path))
            return False
        raise
    try:
        _validate_zip(tmp)
    except Exception:
        _remove_quiet(tmp)
        raise
    if replace_with_retry(tmp, path, attempts, delay, max_delay, log=log):
        return True
    _remove_quiet(tmp)
    return False
