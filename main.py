
import os
import glob
import sys
from pathlib import Path
from shutil import copyfile, copy
import PySimpleGUI as sg
if os.name == 'nt':
    import ctypes
    from ctypes import windll, wintypes
    from uuid import UUID
else:
    from pathlib import Path

from pdf2image import convert_from_path
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


sg.theme('Material1')
index = 0
# 0 not initialized, 1 error, 2 file read, 3 converting, 4 interrupted, 5 finished
status = 0
max_row = 0
prefix_path = ''


if os.name == 'nt':
    class GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", wintypes.DWORD),
            ("Data2", wintypes.WORD),
            ("Data3", wintypes.WORD),
            ("Data4", wintypes.BYTE * 8)
        ] 

        def __init__(self, uuid_):
            ctypes.Structure.__init__(self)
            self.Data1, self.Data2, self.Data3, self.Data4[0], self.Data4[1], rest = uuid_.fields
            for i in range(2, 8):
                self.Data4[i] = rest>>(8 - i - 1)*8 & 0xff


    class UserHandle:
        current = wintypes.HANDLE(0)
        common  = wintypes.HANDLE(-1)


    def get_path(folderid, user_handle=UserHandle.current):
        _CoTaskMemFree = windll.ole32.CoTaskMemFree
        _CoTaskMemFree.restype= None
        _CoTaskMemFree.argtypes = [ctypes.c_void_p]

        _SHGetKnownFolderPath = windll.shell32.SHGetKnownFolderPath
        _SHGetKnownFolderPath.argtypes = [
            ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)
        ] 

        fid = GUID(folderid) 
        pPath = ctypes.c_wchar_p()
        S_OK = 0
        if _SHGetKnownFolderPath(ctypes.byref(fid), 0, user_handle, ctypes.byref(pPath)) != S_OK:
            raise PathNotFoundException()
        path = pPath.value
        _CoTaskMemFree(pPath)
        return path

    doc_id = UUID('{FDD39AD0-238F-46AF-ADB4-6C85480369C7}')
    doc_path = get_path(doc_id)
    prefix_path = doc_path + "\\" + 'renamed'
else:
    doc_path = str(os.path.join(Path.home(), 'Documents'))
    prefix_path = doc_path + "/" + 'renamed'


layout = [ [sg.Text('  源文件夹 ', font=("Helvetica", 14)), 
            sg.Input(key='-file-', enable_events=True, font=("Helvetica", 13), size=(36, 1)), 
            sg.FolderBrowse(button_text='打开文件夹', font=("Helvetica", 14))],
            [sg.Text('输出文件夹', font=("Helvetica", 14)), 
            sg.Input(default_text=prefix_path, size=(36, 1), font=("Helvetica", 13), key='-output-', enable_events=True),
            sg.FolderBrowse(button_text='打开文件夹', font=("Helvetica", 14))],
            [sg.Button('批量重命名PDF', font=("Helvetica", 14)), sg.Button('Exit', font=("Helvetica", 14)),
            sg.StatusBar(text='就绪', key='file_update', font=("Helvetica", 15), size=(10, 1), 
            justification='center', auto_size_text=False, pad=(10, 7)),
            sg.Button('打开生成文件夹', font=("Helvetica", 14), key='-view-', visible=False)],
            [sg.ProgressBar(max_value=10, orientation='h', size=(40, 22), key='progress', visible=False, pad=(10, 1))]]

window = sg.Window('PDF OCR and Auto-renaming System', layout, location=(250, 40))
progress_bar = window['progress']



def ocr_image(pdfs, prefix):
    global status
    global prefix_path
    global index

    status = 3  #working
    progress_bar.update(index, visible=True)
    Path(prefix).mkdir(parents=True, exist_ok=True)
    window['file_update'].update('生成中~~~')

    new_name_list = []

    for pdf in pdfs:
        # file_name = os.path.basename(pdf)
        pages = convert_from_path(pdf_path=pdf, dpi=350)

        # image_name = "Page2_" + os.path.splitext(file_name)[0]+'.jpg'
        # pages[0].save(image_name, "JPEG")
        im = pages[0]
        area1 = (2172, 771, 2498, 905)
        part1 = im.crop(area1)
        area2 = (363, 920, 1880, 1081)
        part2 = im.crop(area2)  

        options1 = "--psm 7 -c tessedit_char_blacklist=\/:*<>|?!"
        options2 = "--psm 6 -c tessedit_char_blacklist=\/:*<>|?!"

        part1_str = str(pytesseract.image_to_string(part1, config=options1)).strip()
        part2_str = str(pytesseract.image_to_string(part2, config=options2)).partition('\n')[0].strip()
        new_name_list.append(part1_str + ' - ' + part2_str + '.pdf')
        # print(part1_str)
        # print(part2_str)
        copy(pdf, prefix)
        index += 1
        progress_bar.update_bar(index, max_row)

    os.chdir(prefix)
    for i, pdf in enumerate(pdfs):
        os.rename(pdf, new_name_list[i])
    index += 1

    status = 5
    index = 0


def read_data(infolder):
    global status
    global max_row
    status = 1
    max_row = 0

    os.chdir(infolder)
    pdf_data = []
    for file in glob.glob("*.pdf"):
        print(file)
        pdf_data.append(file)
    max_row = len(pdf_data) + 1
    if max_row ==0:
         sg.popup_error('没有发现PDF文件')
    else:
        status = 2
    return pdf_data



def make_path(prefix, path):
    return prefix + "\\" + os.path.basename(path)


def main():
    prefix = None
    global status
    global prefix_path
    global max_row
    pdf_data = {}
    folder_imported = False

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == '-file-':
            window['file_update'].update('文件地址已输入')
            pdf_data = read_data(values['-file-'])
            if status == 2:
                if folder_imported:
                    window['file_update'].update('数据已经更新')
                else:
                    folder_imported = True
                    window['file_update'].update('数据已经导入')

        if event == '-output-':
            prefix = values['-output-']
        if event == '批量重命名PDF':
            if status == 2 and pdf_data:
                if prefix is None:
                    ocr_image(pdf_data, prefix_path)
                else:
                    ocr_image(pdf_data, str(prefix))
                    prefix_path = str(prefix)
        if event == '-view-':
            if prefix_path:
                if os.name == 'nt':
                    os.startfile(prefix_path)
                else:
                    opener ="open" if sys.platform == "darwin" else "xdg-open"
                    subprocess.call([opener, prefix_path])

        if status == 5:
            window['file_update'].update('任务完成!')
            window['-view-'].update(visible=True)
            max_row = 0
            status = 0

        elif status == 1:
            window['file_update'].update('未发现PDF文件!')
            max_row = 0
            status = 0

    window.close()


if __name__ == "__main__":
    main()