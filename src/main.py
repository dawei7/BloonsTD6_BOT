import argparse
import pandas as pd
import time
import pydirectinput
import pyautogui


def play(mode, control_path):
    # Load the Excel data into a DataFrame
    df = pd.read_excel(control_path, sheet_name='playground', skiprows=13, usecols="G:N")
    df_actions = pd.read_excel(control_path, sheet_name='actions')
    dict_actions = df_actions.set_index('action').to_dict()['hotkey']
    
    dict_aliases = {}
    
    for index, row in df.iterrows():
        if pd.notna(row["Alias|Tower|Coordinates"]):
            alias, tower, coordinates = row["Alias|Tower|Coordinates"].split("|")
            dict_aliases[alias] = {
                "tower": tower,
                "coordinates": coordinates
            }
        

    # Function to press a sequence of keys
    def press_hotkey(hotkey):
        keys = hotkey.split('+')
        for key in keys:
            pydirectinput.keyDown(key)
        for key in keys:
            pydirectinput.keyUp(key)

    # Process the rows and execute each action
    white_pixel_count = 0
    first_round = True

    if mode in ["easy|standard", "easy|primary_only", "medium|standard", "medium|military_only", "medium|apopalypse", "medium|reverse]"]:
        round = 1
        max_round = 50
    elif mode in ["hard|standard", "hard|magic_monkeys_only", "hard|double_hp_moabs", "hard|half_cash", "hard|alternate_bloons_rounds"]:
        round = 3
        max_round = 80
    elif mode in ["impoppable", "chimps"]:
        round = 6
        max_round = 100
    elif mode in ["easy|deflation"]:
        round = 20
        max_round = 60
    else:
        raise ValueError(f"Invalid mode: {mode}")

    for index, row in df.iterrows():
        if index < round-1:
            continue
        
        # Wait until the specified second
        if not first_round:
            while True:
                if pyautogui.pixel(970, 720) == (255, 255, 255):
                    round += 1
                    print(round)
                    pyautogui.moveTo(968, 718)
                    pyautogui.click()
                    pyautogui.moveTo(505, 50)
                    pyautogui.click()
                    break
                for hotkey in ["1","2","3","4","5","6","7","8","9","0","-","="]:
                    press_hotkey(hotkey)
                time.sleep(1)

        if round == max_round:
            break
        
        for col_prefix in ["Action1", "Action2", "Action3", "Action4", "Action5", "Action6", "Action7"]:
            if pd.notna(row[col_prefix]):
                
                
                alias, upgrade = row[col_prefix].split("|")
                
                tower = dict_aliases[alias]["tower"]

                upgrade = list(upgrade)
                
                if upgrade[0] == "0" and upgrade[1] == "0" and upgrade[2] == "0":
                    action = tower
                elif upgrade[0] != "x":
                    action = "UpgradePath1"
                elif upgrade[1] != "x":
                    action = "UpgradePath2"
                elif upgrade[2] != "x":
                    action = "UpgradePath3"
                else:
                    raise ValueError(f"Invalid upgrade path: {upgrade}")
                
                coords = dict_aliases[alias]["coordinates"].split(",")
                x, y = map(int, coords)

                # Move to coordinates and click with a random delay
                pyautogui.moveTo(505, 50)
                pyautogui.click()
                pyautogui.moveTo(x, y)
                pyautogui.click()
                
                # Trigger hotkey if mapped,.
                if action == "UpgradePath1":
                    if x<512:
                        pyautogui.moveTo(965, 365)
                        pyautogui.click()
                    else:
                        pyautogui.moveTo(195, 365)
                        pyautogui.click()
                elif action == "UpgradePath2":
                    if x<512:
                        pyautogui.moveTo(965, 465)
                        pyautogui.click()
                    else:
                        pyautogui.moveTo(195, 465)
                        pyautogui.click()
                elif action == "UpgradePath3":
                    if x<512:
                        pyautogui.moveTo(965, 565)
                        pyautogui.click()
                    else:
                        pyautogui.moveTo(195, 565)
                        pyautogui.click()
                    
                else:
                    if action in dict_actions:
                        hotkey = dict_actions[action]
                        press_hotkey(hotkey)
                    pyautogui.click()
                pyautogui.moveTo(505, 50)
                pyautogui.click()
        
        if first_round:
            first_round = False
            pyautogui.moveTo(968, 718)
            pyautogui.click()
            pyautogui.moveTo(968, 718)
            pyautogui.click()
            pyautogui.moveTo(505, 50)
            pyautogui.click()
            
def main():
    parser = argparse.ArgumentParser(description='Process some parameters.')
    parser.add_argument('--mode', type=str, required=True, help='Mode of operation')
    parser.add_argument('--control_path', type=str, required=True, help='Path to the control Excel file')
    args = parser.parse_args()
    
    play(args.mode, args.control_path)
    
if __name__ == '__main__':
    main()

# C:/Users/david/Desktop/bloons6_bot/venv/Scripts/python.exe C:/Users/david/Desktop/bloons6_bot/src/main.py --mode "chimps" --control_path "C:/Users/david/Desktop/bloons6_bot/src/control.xlsm