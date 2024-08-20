from time import sleep
# pip install pypiwin32
import win32com.client
# pip install pyautogui
import pyautogui

# pyinstaller python3.7まで対応 versionに注意
# python -m pip install --upgrade pip
# pip install pyinstaller
# pyinstaller ff14_auto_crafter_gui.py --onefile

def ff14_macro_one(shell, macro_key, decision_key, exec_cnt, sleep_cnt, exec_shutdown, left_key = None):
    print('FF14 Crafter Macro One By Niku Kyuu')
    sleep(1)
    # Ff14 をアクティベート
    shell.AppActivate('FINAL FANTASY XIV')
    # 決定キーを押下
    key_down(shell, decision_key)

    for cnt in range(int(exec_cnt)):
        # 決定キーを押下
        key_down(shell, decision_key)
        # 決定キーを押下
        key_down(shell, decision_key)
        # 決定キーを押下
        key_down(shell, decision_key)
        # マクロキーを押下
        key_down(shell, macro_key)
        # 指定秒数wait
        sleep(int(sleep_cnt))

        if cnt < 10:
            echo_cnt = '/echo ***  ' + str(cnt + 1) + ' / ' + str(exec_cnt) + ' 回終了 ***'
            #print('***  ', cnt + 1, ' / ', exec_cnt, ' 回終了 ***')
        else:
            echo_cnt = '/echo *** ' + str(cnt + 1) + ' / ' + str(exec_cnt) + ' 回終了 ***'
            #print('*** ', cnt + 1, ' / ', exec_cnt, ' 回終了 ***')
        inputStr(shell, echo_cnt)
    # 制作終了
    key_down(shell, 'esc')
    #print('終了')
    inputStr(shell, '/echo FINISHED <se.4>')

    # shutdown 処理
    if exec_shutdown == 'y':
        inputStr(shell, '/shutdown')
        key_down(shell, left_key)
        key_down(shell, decision_key)

def ff14_macro_two(shell, first_key, second_key, decision_key, exec_cnt, sleep_cnt, exec_shutdown, left_key = None):
    print('FF14 Crafter Macro Two By Niku Kyuu')
    sleep(2)
    # Ff14 をアクティベート
    shell.AppActivate('FINAL FANTASY XIV')
    # 決定キーを押下
    #wshell.SendKeys(str(decision_key))
    key_down(shell, decision_key)

    for cnt in range(int(exec_cnt)):
        # 決定キーを押下
        key_down(shell, decision_key)
        # 決定キーを押下
        key_down(shell, decision_key)
        # 決定キーを押下
        key_down(shell, decision_key)
        # 1つ目のマクロキーを押下
        key_down(shell, first_key)
        # 指定秒数wait
        sleep(int(sleep_cnt))
        # 2つ目のマクロキーを押下
        key_down(shell, second_key)
        # 指定秒数wait
        sleep(int(sleep_cnt))

        if cnt < 10:
            echo_cnt = '/echo ***  ' + str(cnt + 1) + ' / ' + str(exec_cnt) + ' 回終了 ***'
            #print('***  ', cnt + 1, ' / ', exec_cnt, ' 回終了 ***')
        else:
            echo_cnt = '/echo *** ' + str(cnt + 1) + ' / ' + str(exec_cnt) + ' 回終了 ***'
            #print('*** ', cnt + 1, ' / ', exec_cnt, ' 回終了 ***')
        inputStr(shell, echo_cnt)
    # 制作終了
    key_down(shell, 'esc')
    #print('終了')
    inputStr(shell, '/echo FINISHED <se.4>')

    # shutdown 処理
    if exec_shutdown == 'y':
        inputStr(shell, '/shutdown')
        key_down(shell, left_key)
        key_down(shell, decision_key)

def key_down(shell, target_key):
    shell.AppActivate('FINAL FANTASY XIV')
    sleep(1)
    pyautogui.press(str(target_key))

    return None

def typewrite(shell, target_key):
    shell.AppActivate('FINAL FANTASY XIV')
    sleep(1)
    pyautogui.typewrite(str(target_key))

    return None

def inputStr(shell, echo_str):
    key_down(shell, 'enter')
    typewrite(shell, echo_str)
    key_down(shell, 'enter')

    return None

if __name__ == '__main__':
    # shellの作成
    shell = win32com.client.Dispatch('WScript.Shell')
    #macro_cnt = input('Enter How Many Macro Palette: ')
    macro_cnt = input('使用するマクロパレットの数を入力: ')
    if int(macro_cnt) == 1:
        #macro_key = input('Enter Macro Key Button: ') # マクロ実行キー
        macro_key = input('実行するマクロのキーボードを入力: ') # マクロ実行キー
        #decision_key = input('Enter Decision Key Button: ') # 決定キー
        decision_key = input('決定キーのキーボードを入力: ') # 決定キー
        #exec_cnt = input('Input Exec cnt: ') # マクロ実行回数
        exec_cnt = input('マクロを実行する回数を入力: ') # マクロ実行回数
        #sleep_cnt = input('Input Time Diff: ') # マクロキー押下後のsleep時間(次の実行までのsleep)
        sleep_cnt = input('マクロ実行後の待機時間(秒)を入力: ') # マクロキー押下後のsleep時間(次の実行までのsleep)
        #exec_shutdown = input('If you want to SHUTDOWN after Craft? y or n: ') # クラフト終了後にシャットダウンするかどうか
        exec_shutdown = input('クラフト終了後にシャットダウンしますか? y or n: ') # クラフト終了後にシャットダウンするかどうか
        if exec_shutdown == 'y':
            #left_key = input('Enter Move Left Side Key: ')
            left_key = input('カーソルを左へ移動するキーボードを入力: ')
            ff14_macro_one(shell, macro_key, decision_key, exec_cnt, sleep_cnt, exec_shutdown, left_key)
        else:
            ff14_macro_one(shell, macro_key, decision_key, exec_cnt, sleep_cnt, exec_shutdown)
    elif int(macro_cnt) == 2:
        #first_key = input('Enter First Macro Key Button: ') # マクロ実行1枚目キー
        first_key = input('1 回目に実行するマクロのキーボードを入力: ') # マクロ実行1枚目キー
        #second_key = input('Enter Second Key Button: ') # マクロ実行2枚目キー
        second_key = input('2 回目に実行するマクロのキーボードを入力: ') # マクロ実行2枚目キー
        #decision_key = input('Enter Decision Key Button: ') # 決定キー
        decision_key = input('決定キーのキーボードを入力: ') # 決定キー
        #exec_cnt = input('Input Exec cnt: ') # マクロ実行回数
        exec_cnt = input('マクロを実行する回数を入力: ') # マクロ実行回数
        #sleep_cnt = input('Input Wait Time: ') # マクロキー押下後のsleep時間(次の実行までのsleep)
        sleep_cnt = input('マクロ実行後の待機時間(秒)を入力: ') # マクロキー押下後のsleep時間(次の実行までのsleep)
        #exec_shutdown = input('If you want to SHUTDOWN after Craft? y or n: ') # クラフト終了後にシャットダウンするかどうか
        exec_shutdown = input('クラフト終了後にシャットダウンしますか? y or n: ') # クラフト終了後にシャットダウンするかどうか
        if exec_shutdown == 'y':
            #left_key = input('Enter Move Left Side Key: ')
            left_key = input('カーソルを左へ移動するキーボードを入力: ')
            ff14_macro_two(shell, first_key, second_key, decision_key, exec_cnt, sleep_cnt, exec_shutdown, left_key)
        else:
            ff14_macro_two(shell, first_key, second_key, decision_key, exec_cnt, sleep_cnt, exec_shutdown)
    else:
        print('マクロパレット2枚以上は対応しておりません。すみません。')
