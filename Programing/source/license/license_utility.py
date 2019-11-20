import json
import re
import os
import datetime
import hashlib
import subprocess


def subprocess_args(include_stdout=True):
    # The following is true only on Windows.
    if hasattr(subprocess, 'STARTUPINFO'):
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        env = os.environ
    else:
        si = None
        env = None
    if include_stdout:
        ret = {'stdout': subprocess.PIPE}
    else:
        ret = {}
    ret.update({'stdin': subprocess.PIPE,
                'stderr': subprocess.PIPE,
                'startupinfo': si,
                'env': env})
    return ret


def read_json(json_str):
    try:
        return json.loads(json_str)
    except:
        return None


def read_uuid():
    output = subprocess.check_output('wmic csproduct get uuid', **subprocess_args(False))
    current_machine_id = output.decode().split('\n')[1].strip()
    current_machine_id = ''.join(current_machine_id.split("-"))
    return ':'.join(re.findall('..', current_machine_id))


def get_running_file_name():
    script_name = os.path.basename(__file__)
    return script_name


def str_2_date(_str):
    try:
        _date = datetime.datetime.strptime(_str, "%Y-%m-%d %H:%M:%S")
    except:
        return None
    return _date


def write_file(filename, content, mode=1, redundancy=""):
    """write encrypted mac to ac_solution.c2v file with redudancy data"""
    try:
        if mode == 1:
            content = content.decode('utf-8') + redundancy
            # remains if x is newline else adds 128
            content = [chr(x) if x != 138 else '\n' for x in list(map(lambda x: ord(x) + 128, content))]
            content = ''.join(content)
            with open(filename, 'w', encoding='utf-8') as filehandle:
                filehandle.write(content)
        else:
            with open(filename, 'w') as filehandle:
                filehandle.write(content)
    except Exception:
        pass


def get_check_path(bot_id):
    if isinstance(bot_id, str):
        _str = read_uuid() + bot_id
    else:
        _str = read_uuid() + str(bot_id)
    if os.name == 'nt':
        return os.path.join('C:/Windows/Temp', hashlib.sha256(_str.encode()).hexdigest())
    else:
        return os.path.join('/tmp', hashlib.sha256(_str.encode()).hexdigest())

