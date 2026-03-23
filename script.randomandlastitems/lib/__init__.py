
import json
import time

import xbmc
import xbmcaddon
from xbmcgui import Window

# Define global variables
RALI_GLOBALS = {'LIMIT': 20,
                 'METHOD': 'Random',
                 'REVERSE': False,
                 'MENU': '',
                 'PLAYLIST': '',
                 'PROPERTY': '',
                 'RESUME': 'False',
                 'SORTBY': '',
                 'TYPE': '',
                 'UNWATCHED': 'False'}
START_TIME: float = time.time()
WINDOW = Window(10000)
MONITOR = xbmc.Monitor()
# Nexus JSON RPC 12.9.0 required for userrating
JSON_RPC_NEXUS: bool = (json.loads(xbmc.executeJSONRPC(
                        '{"jsonrpc": "2.0", "method": "JSONRPC.Version", "id": 1}'))['result']['version']['major'],
                        json.loads(xbmc.executeJSONRPC(
                            '{"jsonrpc": "2.0", "method": "JSONRPC.Version", "id": 1}'))['result']['version']['minor']) >= (12, 9)

addon = xbmcaddon.Addon()
ADDONVERSION = addon.getAddonInfo('version')
ADDONID = addon.getAddonInfo('id')
ADDONNAME = addon.getAddonInfo('name')
