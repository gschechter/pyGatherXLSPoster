import os,configparser, requests, json, openpyxl
import colorama
from colorama import Fore, Back, Style
colorama.init()
print("Welcome to pyGatherXLSPoster")
__location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))

def is_json(myjson):
    try:
        json_object = json.loads(myjson)
    except ValueError as e:
        return False
    return True

config = configparser.ConfigParser()

#If config file doesnt exist,  create with some basic settings
if not (config.read('config.ini')):
    config['Gather.Town Endpoints'] = {'getMap': 'https://gather.town/api/getMap','setMap': 'https://gather.town/api/setMap'}
    config['API Settings'] = {'apiKey': 'abc1234','spaceid': 'abc1234'}
    config['XLS'] = {'Filename': 'posters.xlsx','Worksheet': 'posters'}
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

if not(os.path.exists(os.path.join(__location__, config['XLS']['Filename']))):
    print(f"{Fore.RED} {config['XLS']['Filename']} not found {Style.RESET_ALL} \n\n")
else:


    print(f"{Fore.YELLOW}Please enter apiKey or press leave blank to load config value ({config['API Settings']['apiKey']}):{Style.RESET_ALL}", end='')
    #let user override apikey/spaceid from settings
    apiKey = input("")
    if (apiKey == ""):
        apiKey = config['API Settings']['apiKey']
    print(f"{Fore.YELLOW}Please enter spaceId (ie. dkj63wrer8\space) or press leave blank to load config value ({config['API Settings']['spaceid']}):{Style.RESET_ALL}", end='')
    spaceid = input("")
    if (spaceid == ""):
        spaceid = config['API Settings']['spaceid']


    wb = openpyxl.load_workbook(filename = os.path.join(__location__, config['XLS']['Filename']))
    ws = wb['posters']

    maps = []

    for rows in ws.iter_rows(min_row=2):
        if not any(map['mapid'] == rows[0].value for map in maps):
            map = {'mapid':rows[0].value,'posters':[{'blurb':rows[1].value,'url':rows[2].value}]}
            maps.append(map.copy())
        else:
            for map in maps:
                    if map['mapid'] == rows[0].value:
                        if not any(poster['blurb'] == rows[1].value for poster in map['posters']):
                            poster = {'blurb':rows[1].value,'url':rows[2].value}
                            map['posters'].append(poster.copy())

    print(maps)
    for room in maps:
        print(f"Preparing to getMap {room['mapid']} from GatherTown...")
        geturl = f"{config['Gather.Town Endpoints']['getMap']}?apiKey={apiKey}&spaceId={spaceid}&mapId={room['mapid']}"
        payload={}
        headers = {}
        response = requests.request("GET", geturl, headers=headers, data=payload)

        if is_json(response.text):
            print(f"{Fore.GREEN}Got {Fore.YELLOW}{Style.BRIGHT}{room['mapid']}{Style.RESET_ALL}{Fore.GREEN} from GatherTown begining to search for poster matches from spreadsheet {Style.RESET_ALL}")
            map = json.loads(response.text)
            for object in map['objects']:
                if "properties" in object:
                    if "blurb" in object['properties']:
                        for poster in room["posters"]:
                            if(poster["blurb"] == object['properties']['blurb']):
                                print(f"\tFound match and updating {Fore.YELLOW}{poster['blurb']}{Style.RESET_ALL} @ {Fore.YELLOW}{room['mapid']}{Style.RESET_ALL} with {Fore.BLUE}{poster['url']}{Style.RESET_ALL}")
                                object['properties']['image'] = poster['url']
                                object['properties']['preview'] = poster['url']

            print(f"Sending updated map file for {room['mapid']} to GatherTown...")
            payload=json.dumps({
                "apiKey": apiKey,
                "spaceId": spaceid,
                "mapId": room['mapid'],
                "mapContent": map})
            headers = {'Content-Type': 'application/json'}
            response = requests.request("POST", config['Gather.Town Endpoints']['setMap'], headers=headers, data=payload)
            if(response.text == "done"):
                print(f"{Fore.GREEN}Updated {room['mapid']}{Style.RESET_ALL}")
            else:
                print(f"{Fore.RED}{response}{Style.RESET_ALL}")
                print(f"{Fore.RED}{response.text}{Style.RESET_ALL}")
        else:
            if (response.status_code == 200) and (response.text == ""):
                print(f"{Fore.RED}Response on getMap blank, check if \"{room['mapid']}\" is correct{Style.RESET_ALL}")
            else:
                print(f"{Fore.RED}{response}{Style.RESET_ALL}")
                print(f"{Fore.RED}{response.text}{Style.RESET_ALL}")



#keep window open when done
input("Press enter to continue...")