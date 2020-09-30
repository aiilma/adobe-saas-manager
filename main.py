import os
import psutil
import shutil
import subprocess
import time
import zipfile
from shutil import copyfile
from threading import Timer
from termcolor import cprint

PROCESS_AE = "AfterFX.exe"
PROCESS_PS = "Photoshop.exe"
ACTIVE_ORDERS_PATH = 'C:/artworch_share/active_orders/'
TEMP_PATH = 'C:/artworch_share/tmp/'
DIR_HOME = 'C:/artworch_share/'
DIR_ACTIVE_ORDERS = os.listdir(ACTIVE_ORDERS_PATH)

# сортирует по дате и возвращает самый старый заказ
def get_close_order(dir_actual_orders):
    if len(dir_actual_orders) != 0:
        text_files = finder_info_files(dir_actual_orders)
        full_list = [os.path.join(TEMP_PATH, i) for i in text_files]
        time_sorted_list = sorted(full_list, key=os.path.getmtime)
        return time_sorted_list[0]
    return False

# сортирует по дате и возвращает самый старый заказ
def get_newest_log_fname(dir_logs, DPATH_LOG):
    if len(dir_logs) != 0:
        text_files = finder_info_files(dir_logs)
        full_list = [os.path.join(DPATH_LOG, i) for i in text_files]
        time_sorted_list = sorted(full_list, key=os.path.getmtime)
        return time_sorted_list[len(time_sorted_list)-1]
    return False

# получает абсолютный путь к файлу. возвращает токен заказа
def dir2token(fullpath, lvl=1):
    if lvl == 1:
        file_name = fullpath[22:]
        return file_name[:-4:]
    elif lvl == 0:
        return fullpath[22:]

# поисковик картинок по токену заказа
def finder_by_token(order_token, dir_src_pic=TEMP_PATH + 'pics'):
    '''
    первый аргумент - токен заказа
    второй аргумент - шаблон пути, по которому нужно искать
    '''
    pic_lst = []
    listdir_pic = os.listdir(dir_src_pic)
    for filename in listdir_pic:
        countlet = 0
        for letteri in filename:
            if letteri == '.':
                break
            countlet += 1
        if len(filename[4:countlet]) == len(order_token) and (order_token in filename):
            if filename[-4::] == '.png' or (filename[-4::] == '.jpg' or filename[-5::] == '.jpeg'):
                pic_lst.append(filename)
            else:
                cprint("Picture type undefined!", "red")
                return False
    if len(pic_lst) != 0:
        return pic_lst
    cprint("No custom pictures for " + order_token, "yellow")
    return False

# проверяет запущен ли демон After Effects (усовершенствовать до проверки Photoshop)
def check_proc(process_name):
    for i in psutil.process_iter():
        name = i.name()
        if name == process_name:
            return True
    return False

# генератор директорий заказа
def generate_order_directory(order_dir):
    # формирование папок в директории заказа
    # предполагается, что значение аргумента order_dir заканчивается слэшем
    template_dir_psd_path = order_dir + 'psd/'
    template_dir_total_path = order_dir + 'TOTAL/'
    template_dir_bmat_path = order_dir + 'baseMat/'

    dir_by_token = os.path.dirname(order_dir)
    dir_order_psd = os.path.dirname(template_dir_psd_path)
    dir_order_total = os.path.dirname(template_dir_total_path)
    dir_order_bmat = os.path.dirname(template_dir_bmat_path)

    cprint('Creating directories...', 'yellow')

    os.makedirs(dir_by_token)
    os.makedirs(dir_order_psd)
    os.makedirs(dir_order_total)
    os.makedirs(dir_order_bmat)

    if os.path.exists(dir_by_token) and os.path.exists(dir_order_psd) and os.path.exists(dir_order_total) and os.path.exists(dir_order_bmat):
        cprint(':)', 'green')
        return True
    return False

# инспектор текстовых файлов во временной папке
def finder_info_files(path):
    text_files_lst = []
    for file in path:
        if file[-4::] == '.txt':
            text_files_lst.append(file)
    if len(text_files_lst) != 0:
        return text_files_lst
    return False

# инициализирует работу фотошопа по токену
def ps_init(token):
    # print(check_session(PROCESS_PS))
    if mvlog(token):
        mvpsd(token)
        psd_rename(TEMP_PATH + 'psd/')
        cmd = r'"C:\Program Files\Adobe\Adobe Photoshop CC 2018\Photoshop.exe" "C:\Users\admin\Documents\Adobe Scripts\main_PS.jsx" '
        cprint("Starting " + PROCESS_PS, "green")
        returned_value = subprocess.call(cmd, shell=False)
        print('Returned value: ', returned_value)

# находит количество файлов в директории по расширению файла
def filefinder_byprefix(prefix, place):
    src_dir = os.listdir(place)
    files_count = 0
    prefixlen = len(prefix)
    for filename in src_dir:
        if filename[-prefixlen:] == prefix:
            files_count += 1
    return files_count

# возвращает true, если log успешно перемещен в директорию заказа по токену
def mvlog(token):
    # если есть сохраненные psd в директории заказа
    if filefinder_byprefix('.psd', ACTIVE_ORDERS_PATH + token + '/psd') != 0:
        project_id = 0
        project_type = 0
        comp_flag = 'standard'
        with open(ACTIVE_ORDERS_PATH + token + '/' + token + '.txt', 'r', encoding='utf-8') as udataf:
            flines = udataf.readlines()
            for i, line in enumerate(flines):
                if 'compositions: {' in flines[i]:
                    project_id = flines[i + 1]
                    project_type = flines[i + 2]
                if 'additional' in flines[i]:
                    comp_flag = 'unique'

        proj_num_str = ''
        proj_type_str = ''
        for letter in project_id:
            if letter in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]:
                proj_num_str += letter
        if 'long' in project_type:
            proj_type_str = 'long'
        elif 'short' in project_type:
            proj_type_str = 'short'

        DIR_LOG = DIR_HOME + 'templates/' + comp_flag + '/' + proj_num_str + '/' + proj_type_str + '/main.aep Logs' # шаблон директории проектных ЛОГОВ
        absp_newest_log = get_newest_log_fname(os.listdir(DIR_LOG), DIR_LOG) # самый свежий ЛОГ-файл
        newest_log = absp_newest_log[len(DIR_LOG + '\\'):]
        # парсинг свежего лога и получение пути, куда сохранилась выходная секвенция
        with open(absp_newest_log, 'r', encoding='utf-8') as flog:
            flines = flog.readlines()
            finished_path = ''
            for i, line in enumerate(flines):
                if '  Output To: ' in flines[i]:
                    finished_path = line[len('  Output To: '):]
        # проверка принадлежности токена заказа финишированной секвенции
        if token in finished_path:
            maskdir_psd_log = ACTIVE_ORDERS_PATH + token + '/psd/' + newest_log # шаблон PSD-директории заказчика
            if not os.path.isfile(maskdir_psd_log) and not filefinder_byprefix('.txt', ACTIVE_ORDERS_PATH + token + '/psd'): # если текстового файла ВОВСЕ нет
                os.rename(DIR_LOG + '/' + newest_log, maskdir_psd_log)
                return True
            elif not os.path.isfile(maskdir_psd_log) and filefinder_byprefix('.txt', ACTIVE_ORDERS_PATH + token + '/psd'): # если какой-то текстовый файл уже есть
                print('We should delete this LOG file')
                return True
            else: # если ЛОГ-файл был успешно перемещен в соответствующую директорию заказчика
                return True
    return False

# перемещает psd файлы из директории заказа в сессионную директорию psd по токену
def mvpsd(token):
    tpl_psd_dir = ACTIVE_ORDERS_PATH + token + '/psd/'
    lst_psd_dir = os.listdir(tpl_psd_dir)
    if filefinder_byprefix('.psd', TEMP_PATH + 'psd/') != 0:
        for file in os.listdir(TEMP_PATH + 'psd/'):
            os.remove(TEMP_PATH + 'psd/' + file)

    for file in lst_psd_dir:
        copyfile(tpl_psd_dir + file, TEMP_PATH + 'psd/' + file)

    for file in os.listdir(TEMP_PATH + 'psd/'):
        if file.endswith('.txt'):
            os.remove(TEMP_PATH + 'psd/' + file)

# переименовывает файлы вида [####]_token.psd в [####].psd в директории src_dir
def psd_rename(src_dir):
    def PS_getname(filename):
        countlet = 0
        for letter in filename:
            if letter == '_':
                break
            countlet += 1
        if filename[-4:countlet] != '.psd':
            return filename[:countlet] + '.psd'
        return filename[:countlet]

    for filename in os.listdir(src_dir):
        os.rename(src_dir + filename, src_dir + PS_getname(filename))

# убивает процесс по названию процесса
def process_killer(process_name):
    for proc in psutil.process_iter():
        name = proc.name()
        if name == process_name:
            proc.terminate()
            return True
    return False

# таймер на функцию process_killer
def time_on_procterm(process_name, timearg):
    timeout = time.time() + timearg  # 5 minutes from now
    while True:
        test = 0
        if test == 5 or time.time() > timeout:
            process_killer(process_name)
            break
        test = test - 1

# проверяет, запущена ли сессия
def check_session():
    with open(DIR_HOME + 'session_token.txt', 'r', encoding='utf-8') as fses:
        flines = fses.readlines()
        if len(flines):
            return True
    return False

# поиск лога по токену
def browserLog(token):
    project_id = 0
    project_type = 0
    comp_flag = 'standard'
    with open(ACTIVE_ORDERS_PATH + token + '/' + token + '.txt', 'r', encoding='utf-8') as udataf:
        flines = udataf.readlines()
        for i, line in enumerate(flines):
            if 'compositions: {' in flines[i]:
                project_id = flines[i + 1]
                project_type = flines[i + 2]
            if 'additional' in flines[i]:
                comp_flag = 'unique'

    proj_num_str = ''
    proj_type_str = ''
    for letter in project_id:
        if letter in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]:
            proj_num_str += letter
    if 'long' in project_type:
        proj_type_str = 'long'
    elif 'short' in project_type:
        proj_type_str = 'short'

    #######################################################
    # получение пути, куда сохранилась выходная секвенция #
    #######################################################
    DIR_LOG = DIR_HOME + 'templates/' + comp_flag + '/' + proj_num_str + '/' + proj_type_str + '/main.aep Logs'  # шаблон директории проектных ЛОГОВ
    absp_newest_log = get_newest_log_fname(os.listdir(DIR_LOG), DIR_LOG)  # самый свежий ЛОГ-файл
    if absp_newest_log: # парсинг свежего лога
        with open(absp_newest_log, 'r', encoding='utf-8') as flog:
            flines = flog.readlines()
            finished_path = ''
            for i, line in enumerate(flines):
                if '  Output To: ' in flines[i]:
                    finished_path = line[len('  Output To: '):]
            if token in finished_path:
                return True
    if len(os.listdir(DIR_LOG)):# если в свежаке нет искомого значения, парсить все логи
        for file in os.listdir(DIR_LOG):
            with open(DIR_LOG + '/' + file, 'r', encoding='utf-8') as flog:
                flines = flog.readlines();
                finished_path = ''
                for i, line in enumerate(flines):
                    if '  Output To: ' in flines[i]:
                        finished_path = line[len('  Output To: '):]
                if token in finished_path:
                    return True
    return False

# время, через которое выполнится поиск лога по токену
def time_on_searchlog(token, timearg):
    timeout = time.time() + timearg  # отныне
    while True:
        test = 0
        if test == 5 or time.time() > timeout:
            browserLog(token)
            return True
        test = test - 1

# чистит содержимое текстового файла
def file_clearContent(file):
    with open(file, 'w+', encoding='utf-8') as data:
        data.write('')
        print(data.readlines())
    return False

# время, через которое выполнится поиск длины содержимого папки PSD по токену
def time_on_getlenDPSD(token, timearg):
    timeout = time.time() + timearg  # отныне
    while True:
        test = 0
        if test == 5 or time.time() > timeout:
            udir_psdlen = len(os.listdir(ACTIVE_ORDERS_PATH + token + '/psd'))
            return udir_psdlen
        test = test - 1

# возвращает список файлов по пути path
def getFilesByExt(path, prefix):
    text_files_lst = []
    for file in os.listdir(path):
        if file[-4::] == prefix:
            text_files_lst.append(file)
    if len(text_files_lst) != 0:
        return text_files_lst
    return False

# существует ли архив по токену
def gifzipExists(token):
    ziplst = getFilesByExt(ACTIVE_ORDERS_PATH + token + '/TOTAL/', '.zip')
    if ziplst:
        for fzip in ziplst:
            if token in fzip:
                return True
    return False

# создание и добавление в архив гифок и текстового файла
def gifzip(token):
    if filefinder_byprefix('.gif', ACTIVE_ORDERS_PATH + token + '/TOTAL') != 0:
        giflst = getFilesByExt(ACTIVE_ORDERS_PATH + token + '/TOTAL/', '.gif') # список гифок в директории заказа по токену
        txtlst = getFilesByExt(ACTIVE_ORDERS_PATH + token + '/TOTAL/', '.txt') # список текстовых файлов в директории заказа по токену
        with zipfile.ZipFile(ACTIVE_ORDERS_PATH + token + '/TOTAL/' + token + '.zip', 'w') as myzip:
            for root, dirs, files in os.walk(ACTIVE_ORDERS_PATH + token + '/TOTAL/'):
                for file in giflst:
                    myzip.write(os.path.join(root, file), file, compress_type=zipfile.ZIP_DEFLATED) # запись в объект архива гифок
                for file in txtlst:
                    myzip.write(os.path.join(root, file), file, compress_type=zipfile.ZIP_DEFLATED) # запись в объект архива текстовых файлов
                return True
    return False

# получить данные о пользователе
def getUserData(token):
    tdir_tpl_home = ACTIVE_ORDERS_PATH + token + '/'
    with open(tdir_tpl_home + token + '.txt', encoding='utf-8') as tfile:
        lst = [] # список, хранящий в себе пустые списки
        lst2 = [] # список, хранящий в себе пары ключ-значение
        userData = {}
        for line in tfile:
            line = "".join(line.split())
            line = "\n".join(line.split('info='))
            line = "\n".join(line.split('compositions'))
            line = "\n".join(line.split('{'))
            line = "\n".join(line.split('}'))
            line = "\n".join(line.split(','))
            line = line.split('\n')
            line = line[0]
            line = line.split(':')
            # line[1] = str(line[])
            line = [x for x in line if x] # изоляция от пустых строк в списках
            lst.append(line)
        [lst2.append(i) for i in lst if i != []]
        for pair in lst2:
            userData[pair[0]] = pair[1] # добавление в словарь пар ключ-значение
        return userData

# генератор текстового "spasibo" файла
def spasiboGen(token, version):

    dir_tpl_thxHome = DIR_HOME + 'services/thanks_tpl/'
    tdir_tpl_total = ACTIVE_ORDERS_PATH + token + '/TOTAL/'
    copyfile(dir_tpl_thxHome + 'info_v' + version + '.txt', tdir_tpl_total + 'Hey_ReadME.txt')
    udata = getUserData(token)

    with open(tdir_tpl_total + 'Hey_ReadME.txt', encoding='utf-8') as readme_in:
        content = readme_in.read()

    content = content.replace("<nickname>", udata['usernickname'][1:-1])

    with open(tdir_tpl_total + 'Hey_ReadME.txt', 'w', encoding='utf-8') as readme_out:
        readme_out.write(content)

# перемещает архив из директории по токену во временную папку
def mvgifzip(token, dst):
    src = ACTIVE_ORDERS_PATH + token + '/TOTAL/'
    zipname = token + '.zip'
    if gifzipExists(token):
        copyfile(src + zipname, dst + zipname)
        return True
    return False

# удаляет папку заказа по токену
def dirdel(token):
    try:
        shutil.rmtree(ACTIVE_ORDERS_PATH + token)
    except OSError as e:
        print("Error: %s - %s." % (e.filename, e.strerror))
    if not os.path.exists(ACTIVE_ORDERS_PATH + token):
        return True
    return False


####################################################################################################

# инициализирует работу всех функций
def listener():
    Timer(20, listener).start()
    temp_dir = os.listdir(TEMP_PATH) # директория tmp для token.txt

    if check_session():
        if check_proc(PROCESS_AE):
            cprint("Proccess " + PROCESS_AE + " is running... ", "green")
        elif check_proc(PROCESS_PS):
            cprint("Proccess " + PROCESS_PS + " is running... ", "green")

        if finder_info_files(temp_dir):
            print("Count of orders in queue", len(finder_info_files(temp_dir)))
        else:
            print('No orders in queue :(')
    else: # если сессия неактивна
        cprint("Searching for text files in temporary folder...", "yellow")
        if finder_info_files(temp_dir):  # если в tmp есть заказы
            cprint(":)", "green")

            tmp = get_close_order(temp_dir)  # абсолютный путь до самого ожидающего заказа
            ctoken = dir2token(tmp)

            uroom_path = ACTIVE_ORDERS_PATH + ctoken + '/'  # шаблон комнаты заказа по токену

            if not os.path.exists(uroom_path):
                generate_order_directory(uroom_path)

            if (os.path.exists(uroom_path)) and (len(os.listdir(uroom_path)) == 3): # если папка под заказ существует И ее длина равна 3
                os.rename(TEMP_PATH + dir2token(tmp, 0), uroom_path + dir2token(tmp, 0))  # перемещение текстового файла с информацией о заказе ПО ТОКЕНУ
                if (len(os.listdir(TEMP_PATH + 'pics')) != 0) and (finder_by_token(ctoken) != 0):
                    for picture in finder_by_token(ctoken):
                        os.rename(TEMP_PATH + 'pics/' + picture, ACTIVE_ORDERS_PATH + ctoken + '/baseMat/' + picture)  # перемещение картинок ПО ТОКЕНУ

            with open(DIR_HOME + 'session_token.txt', 'w+', encoding='utf-8') as fsession:
                fsession.write("'" + ctoken + "'")

            cmd = r'"C:\Users\admin\Documents\Adobe Scripts\main_AE.jsx"'
            subprocess.call(cmd, shell=True)
            prefDirlen_upsd = 0
            postDirlen_upsd = 0
            while check_proc('AfterFX.exe'):
                prefDirlen_upsd = len(os.listdir(ACTIVE_ORDERS_PATH + ctoken + '/psd'))
                postDirlen_upsd = time_on_getlenDPSD(ctoken, 5)
                if prefDirlen_upsd == postDirlen_upsd and browserLog(ctoken):
                    process_killer('AfterFX.exe')

            ps_init(ctoken)

            spasiboGen(ctoken, '02') # генерация благодарочки
            if gifzipExists(ctoken) is False: # если архив не существует, то создать его
                gifzip(ctoken)
            if mvgifzip(ctoken, TEMP_PATH + 'archives/'): # если перемещение архива на выдачу в tmp/archives успешно
                dirdel(ctoken) # удаление директории заказа
            file_clearContent(DIR_HOME + 'session_token.txt') # чистка содержимого сессионного файла
        else: # если заказов нет
            cprint("TMP folder is clear", "cyan")


listener()
