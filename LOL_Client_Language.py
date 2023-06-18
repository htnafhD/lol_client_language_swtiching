import os
import argparse
import pythoncom
import win32com.client
import psutil
import glob
import subprocess

array_list_lan = [
    'vi_VN: Vietnamese',
    'ja_JP: Japanese',
    'ko_KR: Korean',
    'zh_CN: Chinese',
    'zh_TW: Taipei',
    'es_ES: Spanish',
    'es_MX: Spanish (Latin America)',
    'en_GR: English - UK',
    'en_AU: English - Australia',
    'fr_FR: French',
    'de_DE: German',
    'it_IT: Italian',
    'pl_PL: Polish',
    'ro_RO: Romanian',
    'el_GR: Greek',
    'pt_BR: Portuguese',
    'hu_HU: Hungarian',
    'ru_RU: Russian',
    'tr_TR: Turkish'
]

def handle_locale():
    lst_f_lan = []
    for each_lan in array_list_lan:
        each_lan = each_lan.split(':')[0]
        lst_f_lan.append(each_lan)
    return lst_f_lan

def get_lan(locale):
    similar_locale = list(filter(lambda x: locale in x, array_list_lan))
    return similar_locale[0].split(':')[1].strip()

def get_disk_name():
    partitions = psutil.disk_partitions()
    disk_names = [partition.device for partition in partitions]
    return disk_names

def search_for_LOL():
    disk_names = get_disk_name()
    lol_client = "LeagueClient.exe"
    for disk in disk_names:
        search_path = os.path.join(f"{disk}", "**", lol_client)
        matching_files = glob.glob(search_path, recursive=True)
        if len(matching_files) == 1:
            return matching_files[0]
        elif len(matching_files) > 1:
            print('[*] Found more than 1 LeagueClient.exe. Please uninstall one!')
            return None
    print('[*] Unable to find LeagueClient.exe')
    return None

def get_the_correct_locale(locale):
    correct_locale = handle_locale()
    if locale in correct_locale:
        return locale
    else:
        similar_locale = list(filter(lambda x: locale in x, correct_locale))
        try:
            user_select = input(f'[*] You mean {similar_locale[0]}? (Y/N)? ').lower()
        except:
            user_select = None
            print('[*] Please -h to view correct locale!')
        if user_select in ['yes', 'y', '']:
            del_existed_ink()
            print('[*] Creating shortcut')
            return similar_locale[0]
        elif user_select in ['no', 'n']:
            print('[*] Please enter the correct locale again!')
            return None
        else:
            print("[*] Allowed input with ['yes', 'y', 'n', 'no'] only!")

def create_shortcut(target, shortcut_name, shortcut_path, arguments):
    shell = win32com.client.Dispatch("WScript.Shell")
    desktop_path = shell.SpecialFolders("Desktop")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target
    shortcut.Arguments = arguments
    shortcut.Description = shortcut_name
    shortcut.Save()

def del_existed_ink():
    try:
        desktop_path = f"{os.environ['USERPROFILE']}\\Desktop\\"
        old_league = list(filter(lambda x: "League of Legends.lnk" in x, os.listdir(desktop_path)))
        old_league_dir = f"{desktop_path}\\{old_league[0]}"
        os.remove(old_league_dir)
        print('[*] Removed old league file!')
    except Exception as e:
        print(f'[*] Can\'t delete file. Exception {e}')

def shortcut(locale):
    f_locale = get_the_correct_locale(locale)
    if f_locale:
        lol_client = search_for_LOL()
        if lol_client:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            arguments  = f'"{lol_client}" --locale={f_locale}'  
            locale_name = f_locale.split('_')[1]
            shortcut_name = f"[{locale_name}] League of Legends" 
            shortcut_path = f"{desktop_path}/{shortcut_name}.lnk"  
            try:
                pythoncom.CoInitialize()
                create_shortcut(lol_client, shortcut_name, shortcut_path, arguments)
                pythoncom.CoUninitialize()
                try:
                    lan = get_lan(f_locale)
                except:
                    lan = f_locale
                print(f'[*] Finished creating shorcut for language {lan} at {shortcut_path}')
                open_shortcut(shortcut_path)
            except Exception as e:
                print(f"[*] Unknow Error {e}")
        else:
            print('[*] Exiting the script!')
    else:
        print('[*] Exiting the script!')

def open_shortcut(shortcut_path):
    print(f'[*] Opening the shortcut file in {shortcut_path}')
    with open(os.devnull, 'w') as devnull:
        subprocess.Popen([shortcut_path], stdout=devnull, stderr=devnull, shell=True)
    print('[*] The language package will take time to download and the speed will depends on your network! Also, please re-open your Client to take effect if it still running.')

def welcome():
    x = '''[*] Support languages:
    'vi_VN: Vietnamese',
    'ja_JP: Japanese',
    'ko_KR: Korean',
    'zh_CN: Chinese',
    'zh_TW: Taipei',
    'es_ES: Spanish',
    'es_MX: Spanish (Latin America)',
    'en_GR: English - UK',
    'en_AU: English - Australia',
    'fr_FR: French',
    'de_DE: German',
    'it_IT: Italian',
    'pl_PL: Polish',
    'ro_RO: Romanian',
    'el_GR: Greek',
    'pt_BR: Portuguese',
    'hu_HU: Hungarian',
    'ru_RU: Russian',
    'tr_TR: Turkish'''
    return x

if __name__ == "__main__":
    help_detail = welcome()
    parser = argparse.ArgumentParser("League Of Legends - Language Switching")
    parser.add_argument("-l", "--locale",  help=f"Locale to switch to.{help_detail}", required=True)
    args = parser.parse_args()
    locale = args.locale
    shortcut(locale)