#!/usr/bin/env python
import os
import re
import shutil
import sqlite3
import sys
import time
import json
import uuid
import zipfile
import collections
from builtins import OSError, input

from pathlib import Path

# You can adjust some setting here. Default is for QOwnNotes app.
media_dir_name = 'media'  # name of the directory inside the produced directory where all images and attachments will be stored
wiz_home_path = Path('C:\\Users\\demo\\Documents\\My Knowledge\\Data\\admin@wiz.cn')
wiz_index_db = Path(wiz_home_path / 'index.db')
creation_date_in_filename = False  # True to insert note creation time to the note file name, False to disable

############################################################################

Notebook = collections.namedtuple('Notebook', ['path', 'media_path'])

WizNote = collections.namedtuple("WizNote", ['path'])


def zip_ya(startdir):
    startdir = str(startdir)
    file_news = startdir + '.ziw'  # 压缩后文件夹的名字
    z = zipfile.ZipFile(file_news, 'w', zipfile.ZIP_DEFLATED)  # 参数一：文件夹名
    for dirpath, dirnames, filenames in os.walk(startdir):
        fpath = dirpath.replace(startdir, '')  # 这一句很重要，不replace的话，就从根目录开始复制
        fpath = fpath and fpath + os.sep or ''  # 这句话理解我也点郁闷，实现当前文件夹以及包含的所有文件的压缩
        for filename in filenames:
            z.write(os.path.join(dirpath, filename), fpath + filename)
            print('压缩成功')
    z.close()


def sanitise_path_string(path_str):
    for char in (':', '/', '\\', '|'):
        path_str = path_str.replace(char, '-')
    for char in ('?', '*'):
        path_str = path_str.replace(char, '')
    path_str = path_str.replace('<', '(')
    path_str = path_str.replace('>', ')')
    path_str = path_str.replace('"', "")
    path_str = path_str.replace("'", "")
    path_str = path_str.replace('\n', "")

    return path_str[:240].strip()


work_path = Path.cwd()
media_dir_name = sanitise_path_string(media_dir_name)

if len(sys.argv) > 1:
    files_to_convert = [Path(path) for path in sys.argv[1:]]
else:
    files_to_convert = Path(work_path).glob('*.nsx')

if not files_to_convert:
    print('No .nsx files found')
    exit(1)

for file in files_to_convert:
    nsx_file = zipfile.ZipFile(str(file))
    config_data = json.loads(nsx_file.read('config.json').decode('utf-8'))
    notebook_id_to_path_index = {}
    wiz_notebook_id_to_path_index = {}

    recycle_bin_path = work_path / Path('Recycle bin')

    n = 1
    while recycle_bin_path.is_dir():
        recycle_bin_path = work_path / Path('{}_{}'.format('Recycle bin', n))
        n += 1

    recycle_bin_media_path = recycle_bin_path / media_dir_name
    recycle_bin_media_path.mkdir(parents=True)
    notebook_id_to_path_index['1027_#00000000'] = Notebook(recycle_bin_path, recycle_bin_media_path)

    print('Extracting notes from "{}"'.format(file.name))
    for notebook_id in config_data['notebook']:
        notebook_data = json.loads(nsx_file.read(notebook_id).decode('utf-8'))
        notebook_title = notebook_data['title'] or 'Untitled'
        notebook_path = work_path / Path(sanitise_path_string(notebook_title))

        n = 1
        while notebook_path.is_dir():
            notebook_path = work_path / Path('{}_{}'.format(sanitise_path_string(notebook_title), n))
            n += 1

        notebook_media_path = Path(notebook_path / media_dir_name)
        notebook_media_path.mkdir(parents=True)

        notebook_id_to_path_index[notebook_id] = Notebook(notebook_path, notebook_media_path)

        wiz_notebook_path = Path(wiz_home_path / 'dsnote' / sanitise_path_string(notebook_title))
        wiz_notebook_path.mkdir(parents=True)
        wiz_notebook_id_to_path_index[notebook_id] = WizNote(wiz_notebook_path)

    note_id_to_title_index = {}
    wiz_note_id_to_title_index = {}
    converted_note_ids = []

    for note_id in config_data['note']:
        note_data = json.loads(nsx_file.read(note_id).decode('utf-8'))

        note_title = note_data.get('title', 'Untitled')
        note_ctime = note_data.get('ctime', '')
        note_mtime = note_data.get('mtime', '')
        note_parent_notebook_id = note_data.get('parent_id','')
        notebook_data = json.loads(nsx_file.read(note_parent_notebook_id).decode('utf-8'))
        notebook_title = notebook_data['title'] or 'Untitled'

        note_id_to_title_index[note_id] = note_title
        wiz_note_id_to_title_index[note_id] = note_title

        try:
            parent_notebook_id = note_data['parent_id']
            parent_notebook = notebook_id_to_path_index[parent_notebook_id]
            wiz_parent_notebook = wiz_notebook_id_to_path_index[parent_notebook_id]
            wiz_note_path = wiz_parent_notebook.path / sanitise_path_string(note_title)
        except KeyError:
            continue

        print('Converting note "{}"'.format(note_title))

        content = re.sub('<img class="[^"]*syno-notestation-image-object" src=[^>]*ref="',
                         '<img src="', note_data.get('content', ''))
        wiz_content = re.sub('<img class="[^"]*syno-notestation-image-object" src=[^>]*ref="',
                             '<img src="', note_data.get('content', ''))

        attachments_data = note_data.get('attachment')
        attachment_list = []
        wiz_attachment_list = []

        if attachments_data:
            for attachment_id in note_data.get('attachment', ''):

                ref = note_data['attachment'][attachment_id].get('ref', '')
                md5 = note_data['attachment'][attachment_id]['md5']
                source = note_data['attachment'][attachment_id].get('source', '')
                name = sanitise_path_string(note_data['attachment'][attachment_id]['name'])

                n = 1
                while Path(parent_notebook.media_path / name).is_file():
                    name_parts = name.rpartition('.')
                    name = ''.join((name_parts[0], '_{}'.format(n), name_parts[1], name_parts[2]))
                    n += 1

                link_path_str = '{}/{}'.format(media_dir_name, name)
                wiz_index_files_link_path_str = '{}/{}'.format('index_files', name)

                html_link_template = '<a href="{}">{}</a>'

                try:
                    Path(parent_notebook.media_path / name).write_bytes(nsx_file.read('file_' + md5))

                    if not Path(wiz_note_path / 'index_files').exists():
                        Path(wiz_note_path / 'index_files').mkdir(mode=0o777, parents=True)
                    Path(wiz_note_path / 'index_files' / name).write_bytes(
                        nsx_file.read('file_' + md5))

                    attachment_list.append(html_link_template.format(link_path_str, name))
                    wiz_attachment_list.append(html_link_template.format(wiz_index_files_link_path_str, name))
                except Exception:
                    if source:
                        attachment_list.append(html_link_template.format(source, name))
                    else:
                        print('Can\'t find attachment "{}" of note "{}"'.format(name, note_title))
                        attachment_list.append(html_link_template.format(source, 'NOT FOUND'))

                if ref and source:
                    content = content.replace(ref, source)
                elif ref:
                    content = content.replace(ref, link_path_str)

                if ref and source:
                    wiz_content = wiz_content.replace(ref, source)
                elif ref:
                    wiz_content = wiz_content.replace(ref, wiz_index_files_link_path_str)

        if attachment_list:
            content = 'Attachments: {}  \n{}'.format(', '.join(attachment_list), content)
        if wiz_attachment_list:
            wiz_content = '<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> Attachments: {}  \n{}'.format(', '.join(wiz_attachment_list), wiz_content)
        if creation_date_in_filename and note_ctime:
            note_title = time.strftime('%Y-%m-%d ', time.localtime(note_ctime)) + note_title

        md_file_name = sanitise_path_string(note_title) or 'Untitled'
        md_file_path = Path(parent_notebook.path / '{}.{}'.format(md_file_name, 'htm'))
        wiz_md_file_path = Path(wiz_note_path / '{}.{}'.format("index", 'html'))

        n = 1
        while md_file_path.is_file():
            md_file_path = Path(parent_notebook.path / ('{}_{}.{}'.format(
                sanitise_path_string(note_title), n, 'htm')))
            n += 1

        if not Path(wiz_note_path).exists():
            Path(wiz_note_path).mkdir(mode=0o777, parents=True)

        md_file_path.write_text(content, 'utf-8')
        wiz_md_file_path.write_text(wiz_content, 'utf-8')

        zip_ya(wiz_note_path)
        shutil.rmtree(wiz_note_path)

        conn = sqlite3.connect(wiz_index_db)
        c = conn.cursor()
        print
        "Opened database successfully";

        sql = 'INSERT INTO "main"."WIZ_DOCUMENT" (' \
              '"DOCUMENT_GUID", "DOCUMENT_TITLE", "DOCUMENT_LOCATION", ' \
              '"DOCUMENT_NAME", "DOCUMENT_SEO", "DOCUMENT_URL", ' \
              '"DOCUMENT_AUTHOR", "DOCUMENT_KEYWORDS", "DOCUMENT_TYPE", ' \
              '"DOCUMENT_OWNER", "DOCUMENT_FILE_TYPE", "STYLE_GUID", ' \
              '"DT_CREATED", "DT_MODIFIED", "DT_ACCESSED", ' \
              '"DOCUMENT_ICON_INDEX", "DOCUMENT_SYNC", "DOCUMENT_PROTECT",' \
              ' "DOCUMENT_READ_COUNT", "DOCUMENT_ATTACHEMENT_COUNT", "DOCUMENT_INDEXED",' \
              ' "DT_INFO_MODIFIED", "DOCUMENT_INFO_MD5", "DT_DATA_MODIFIED",' \
              ' "DOCUMENT_DATA_MD5", "DT_PARAM_MODIFIED", "DOCUMENT_PARAM_MD5",' \
              ' "WIZ_VERSION", "KB_GUID", "WIZ_DOWNLOADED",' \
              ' "WIZ_SERVER_VERSION", "WIZ_LOCAL_FLAGS", "DOCUMENT_SOURCELOCATION",' \
              ' "DATA_CHANGED")  \
                      VALUES (' \
              '"{}","{}","{}",' \
              '"{}",{},{},' \
              '{},{},"{}",' \
              '"{}",{},{},' \
              '"{}","{}","{}",' \
              '{},{},{},' \
              '{},{},{},' \
              '"{}",{},"{}",' \
              '"{}","{}",{},' \
              '{},{},{},' \
              '{},{},{},' \
              '{});'

        sql_format = sql.format(uuid.uuid1(), sanitise_path_string(note_title) ,'/dsnote/' + notebook_title + '/',
                                sanitise_path_string(note_title) + '.ziw', "NULL", "NULL", "NULL", "NULL", "document",
                                "admin@wiz.cn", "NULL",
                                "NULL", note_ctime, note_mtime, note_mtime, -1, 0, 0, 6, 0, 1, note_ctime, "NULL",
                                note_mtime, note_id, note_ctime, "NULL", 2308, "NULL", 1, 2308, "NULL", "NULL", 0, sql)
        print(sql_format)
        conn.execute(sql_format)

        conn.commit()
        print
        "Records created successfully";
        conn.close()

        converted_note_ids.append(note_id)

    for notebook in notebook_id_to_path_index.values():
        try:
            notebook.media_path.rmdir()
        except OSError:
            pass

    not_converted_note_ids = set(note_id_to_title_index.keys()) - set(converted_note_ids)

    if not_converted_note_ids:
        print('Failed to convert notes:',
              '\n'.join(('    {} (ID: {})'.format(note_id_to_title_index[note_id], note_id)
                         for note_id in not_converted_note_ids)),
              sep='\n')

    if len(config_data['notebook']) == 1:
        notebook_log_str = 'notebook'
    else:
        notebook_log_str = 'notebooks'

    print('Converted {} {} and {} out of {} notes.\n'.format(len(config_data['notebook']),
                                                             notebook_log_str,
                                                             len(converted_note_ids),
                                                             len(note_id_to_title_index.keys())))
    try:
        recycle_bin_media_path.rmdir()
        recycle_bin_path.rmdir()
    except OSError:
        pass

input('Press Enter to quit...')
