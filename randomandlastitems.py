# This program is Free Software see LICENSE file for details
""" Skins use to provide random or recent added library info for widgets

Invoke with RunScript() built-in command.  Script will return library info
via Window Properties.  See README.txt for more.

Typical usage example:

    On Home window:

    <onload>RunScript(script.randomandlastitems,limit=12,method=Last,
    playlist="some playlist")</onload>

    This will get library info for the 12 newest (date added) playlist itmes
    and return as window properties.  It runs as a one-shot (not a service)

    Does not provide results for artist or mixed smart playlists
"""


import json
import os
import random
import re
import sys
import time
import urllib.request
from operator import itemgetter
from typing import List, Tuple
from xml.dom.minidom import parse

import xbmc
import xbmcaddon
import xbmcgui
import xbmcvfs
from xbmcgui import Window

# Define global variables
LIMIT = 20
METHOD = 'Random'
REVERSE = False
MENU = ''
PLAYLIST = ''
PROPERTY = ''
RESUME = 'False'
SORTBY = ''
START_TIME: float = time.time()
TYPE = ''
UNWATCHED = 'False'
WINDOW = xbmcgui.Window(10000)
MONITOR = xbmc.Monitor()
# Nexus JSON RPC 12.9.0 required for userrating
JSON_RPC_NEXUS: bool = (json.loads(xbmc.executeJSONRPC(
                        '{"jsonrpc": "2.0", "method": "JSONRPC.Version", "id": 1}'))['result']['version']['major'],
                        json.loads(xbmc.executeJSONRPC(
                            '{"jsonrpc": "2.0", "method": "JSONRPC.Version", "id": 1}'))['result']['version']['minor']) >= (12, 9)

__addon__ = xbmcaddon.Addon()
__addonversion__ = __addon__.getAddonInfo('version')
__addonid__ = __addon__.getAddonInfo('id')
__addonname__ = __addon__.getAddonInfo('name')


def log(txt: str) -> None:
    """utility writes info to Kodi debug level log 

    Args:
        txt (str): text to log

    Returns: None
    """
    message = '{}: {}'.format(__addonname__, txt)
    xbmc.log(msg=message, level=xbmc.LOGDEBUG)


def _getPlaylistType() -> None:
    """sets global variables for a playlist

        Returns:  None
    """
    global METHOD
    global PLAYLIST
    global REVERSE
    global SORTBY
    global TYPE
    _doc = parse(xbmcvfs.translatePath(PLAYLIST))
    _type = _doc.getElementsByTagName('smartplaylist')[0].attributes.item(0).value
    if _type == 'movies':
        TYPE = 'Movie'
    if _type == 'musicvideos':
        TYPE = 'MusicVideo'
    if _type == 'episodes' or _type == 'tvshows':
        TYPE = 'Episode'
    if _type == 'songs' or _type == 'albums':
        TYPE = 'Music'
    if _type == "artists" or _type == 'mixed':
        TYPE = 'Invalid'
    # get playlist name
    _name = ''
    if _doc.getElementsByTagName('name'):
        try:
            _name = _doc.getElementsByTagName('name')[0].firstChild.nodeValue
        except:
            _name = ''
    _setProperty('{}.Name'.format(PROPERTY), str(_name))
    # get playlist order
    if METHOD == 'Playlist':
        if _doc.getElementsByTagName('order'):
            SORTBY = _doc.getElementsByTagName('order')[0].firstChild.nodeValue
            if _doc.getElementsByTagName('order')[0].attributes.item(0).value == 'descending':
                REVERSE = True
        else:
            METHOD = ''


def _timeTook(t: float) -> str:
    """ Utility gets elapsed time for query (used for logging)

    Args:
        t (float): start time

    Returns:
        str: elapsed time to .001 sec
    """
    t = (time.time() - t)
    if t >= 60:
        return '%.3fm' % (t / 60.0)
    return '%.3fs' % (t)


def _watchedOrResume(_total: int, _watched: int, _unwatched: int, _result: list,
                     _file: dict) -> Tuple[int, int, int, list]:
    """Gets watched / in progess status for a library item and increments counters

    RESUME and UNWATCHED are bools to determine when an item is valid for inclusion
    in the _result item list.  eg, if UNWATCHED is true any watched item is
    excluded.

    Args:
        _total (int): cumulative number of items
        _watched (int): cumulative number of watched items
        _unwatched (int): cumulative number of unwatched items
        _result (list): cumulative list of processed library items
        _file (dict): a library item to evaluate watched / in progress status

    Returns:
        Tuple[int, int, int, list]: updated totals and item list
    """
    global RESUME
    global UNWATCHED
    _total += 1
    _playcount: int = _file['playcount']
    _resume: int = _file['resume']['position']
    # Add Watched flag and counter for episodes
    if _playcount == 0:
        _file['watched'] = 'False'
        _unwatched += 1
    else:
        _file['watched'] = 'True'
        _watched += 1
    if (UNWATCHED == 'False' and RESUME == 'False') or (UNWATCHED == 'True' and _playcount == 0) or (RESUME == 'True' and _resume != 0) and _file.get('dateadded'):
        _result.append(_file)
    return _total, _watched, _unwatched, _result


def _getMovies() -> None:
    """retrieves movie info from Kodi library and sets properties

    If a movie playlist is not provided uses movie titles library node

    Returns:
        None
    """
    global LIMIT
    global METHOD
    global MENU
    global PLAYLIST
    global PROPERTY
    global RESUME
    global REVERSE
    global SORTBY
    global UNWATCHED
    _result: List[dict] = []
    _total = 0
    _unwatched = 0
    _watched = 0
    # Request database using JSON
    if PLAYLIST == '':
        PLAYLIST = 'videodb://movies/titles/'
    if JSON_RPC_NEXUS:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "Files.GetDirectory", '
            '"params": {"directory": "%s", '
                '"media": "video", '
                '"properties": '
                    '["title", '
                    '"originaltitle", '
                    '"playcount", '
                    '"year", '
                    '"genre", '
                    '"studio", '
                    '"country", '
                    '"tagline", '
                    '"plot", '
                    '"runtime", '
                    '"file", '
                    '"plotoutline", '
                    '"lastplayed", '
                    '"trailer", '
                    '"rating", '
                    '"userrating", '
                    '"resume", '
                    '"art", '
                    '"streamdetails", '
                    '"mpaa", '
                    '"director", '
                    '"dateadded"]'
                '}, '
            '"id": 1}' % (PLAYLIST))
    else:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "Files.GetDirectory", '
            '"params": {"directory": "%s", '
                '"media": "video", '
                '"properties": '
                    '["title", '
                    '"originaltitle", '
                    '"playcount", '
                    '"year", '
                    '"genre", '
                    '"studio", '
                    '"country", '
                    '"tagline", '
                    '"plot", '
                    '"runtime", '
                    '"file", '
                    '"plotoutline", '
                    '"lastplayed", '
                    '"trailer", '
                    '"rating", '
                    '"resume", '
                    '"art", '
                    '"streamdetails", '
                    '"mpaa", '
                    '"director", '
                    '"dateadded"]'
                '}, '
            '"id": 1}' % (PLAYLIST))
    _json_pl_response: dict = json.loads(_json_query)
    # If request return some results
    _files: dict = _json_pl_response.get('result', {}).get('files')
    if _files:
        for _item in _files:
            if MONITOR.abortRequested():
                break
            if _item['filetype'] == 'directory':
                if JSON_RPC_NEXUS:
                    _json_query = xbmc.executeJSONRPC(
                        '{"jsonrpc": "2.0", '
                        '"method": "Files.GetDirectory", '
                            '"params": '
                                '{"directory": "%s", '
                                '"media": "video", '
                                '"properties": '
                                    '["title", '
                                    '"originaltitle", '
                                    '"playcount", '
                                    '"year", '
                                    '"genre", '
                                    '"studio", '
                                    '"country", '
                                    '"tagline", '
                                    '"plot", '
                                    '"runtime", '
                                    '"file", '
                                    '"plotoutline", '
                                    '"lastplayed", '
                                    '"trailer", '
                                    '"rating", '
                                    '"userrating", '
                                    '"resume", '
                                    '"art", '
                                    '"streamdetails", '
                                    '"mpaa", '
                                    '"director", '
                                    '"dateadded"]'
                                '}, '
                            '"id": 1}' % (_item['file']))
                else:
                    _json_query = xbmc.executeJSONRPC(
                        '{"jsonrpc": "2.0", '
                        '"method": "Files.GetDirectory", '
                            '"params": '
                                '{"directory": "%s", '
                                '"media": "video", '
                                '"properties": '
                                    '["title", '
                                    '"originaltitle", '
                                    '"playcount", '
                                    '"year", '
                                    '"genre", '
                                    '"studio", '
                                    '"country", '
                                    '"tagline", '
                                    '"plot", '
                                    '"runtime", '
                                    '"file", '
                                    '"plotoutline", '
                                    '"lastplayed", '
                                    '"trailer", '
                                    '"rating", '
                                    '"resume", '
                                    '"art", '
                                    '"streamdetails", '
                                    '"mpaa", '
                                    '"director", '
                                    '"dateadded"]'
                                '}, '
                            '"id": 1}' % (_item['file']))
                _json_set_response: dict = json.loads(_json_query)
                _movies: List[dict] = _json_set_response.get(
                    'result', {}).get('files') or []
                if not _movies:
                    log('## MOVIESET {} COULD NOT BE LOADED ##'.format(
                        _item['file']))
                    log('JSON RESULT {}'.format(_json_set_response))
                for _movie in _movies:
                    if MONITOR.abortRequested():
                        break
                    _playcount: int = _movie['playcount']
                    if RESUME == 'True':
                        _resume: int = _movie['resume']['position']
                    else:
                        _resume = 0
                    _total += 1
                    if _playcount == 0:
                        _unwatched += 1
                    else:
                        _watched += 1
                    if (UNWATCHED == 'False' and RESUME == 'False') or (UNWATCHED == 'True' and _playcount == 0) or (RESUME == 'True' and _resume != 0):
                        _result.append(_movie)
            else:
                _playcount = _item['playcount']
                if RESUME == 'True':
                    _resume = _item['resume']['position']
                else:
                    _resume = 0
                _total += 1
                if _playcount == 0:
                    _unwatched += 1
                else:
                    _watched += 1
                if (UNWATCHED == 'False' and RESUME == 'False') or (UNWATCHED == 'True' and _playcount == 0) or (RESUME == 'True' and _resume != 0):
                    _result.append(_item)
        _setVideoProperties(_total, _watched, _unwatched)
        _count = 0
        if METHOD == 'Last':
            _result = sorted(_result, key=itemgetter(
                'dateadded'), reverse=True)
        elif METHOD == 'Playlist':
            _result = sorted(_result, key=itemgetter(SORTBY), reverse=REVERSE)
        else:
            random.shuffle(_result, random.random)
        for _movie in _result:
            if MONITOR.abortRequested() or _count == LIMIT:
                break
            _count += 1
            _json_query = xbmc.executeJSONRPC(
                '{"jsonrpc": "2.0", '
                    '"method": "VideoLibrary.GetMovieDetails", '
                    '"params": '
                        '{"properties": ["streamdetails"], "movieid":%s }, '
                    '"id": 1}' % (_movie['id']))
            _json_query = json.loads(_json_query)
            if 'result' in _json_query and 'moviedetails' in _json_query['result']:
                item = _json_query['result']['moviedetails']
                _movie['streamdetails'] = item['streamdetails']
            if _movie['resume']['position'] > 0 and float(_movie['resume']['total']) > 0:
                resume = 'true'
                played = '%s%%' % int(
                    (float(_movie['resume']['position']) / float(_movie['resume']['total'])) * 100)
            else:
                resume = 'false'
                played = '0%'
            if _movie['playcount'] >= 1:
                watched = 'true'
            else:
                watched = 'false'
            path = media_path(_movie['file'])
            play = 'RunScript(' + __addonid__ + ',movieid=' + (
                str(_movie.get('id')) + ')')
            art = _movie['art']
            streaminfo = media_streamdetails(_movie['file'].lower(),
                                             _movie['streamdetails'])
            # Get runtime from streamdetails or from NFO
            if streaminfo['duration'] != 0:
                runtime = str(int((streaminfo['duration'] / 60) + 0.5))
            else:
                if isinstance(_movie['runtime'], int):
                    runtime = str(int((_movie['runtime'] / 60) + 0.5))
                else:
                    runtime = _movie['runtime']
            # Set window properties
            _setProperty('%s.%d.DBID'            % (PROPERTY, _count), str(_movie.get('id','')))
            _setProperty('%s.%d.Title'           % (PROPERTY, _count), _movie.get('title',''))
            _setProperty('%s.%d.OriginalTitle'   % (PROPERTY, _count), _movie.get('originaltitle',''))
            _setProperty('%s.%d.Year'            % (PROPERTY, _count), str(_movie.get('year','')))
            _setProperty('%s.%d.Genre'           % (PROPERTY, _count), ' / '.join(_movie.get('genre','')))
            _setProperty('%s.%d.Studio'          % (PROPERTY, _count), ' / '.join(_movie.get('studio','')))
            _setProperty('%s.%d.Country'         % (PROPERTY, _count), ' / '.join(_movie.get('country','')))
            _setProperty('%s.%d.Plot'            % (PROPERTY, _count), _movie.get('plot',''))
            _setProperty('%s.%d.PlotOutline'     % (PROPERTY, _count), _movie.get('plotoutline',''))
            _setProperty('%s.%d.Tagline'         % (PROPERTY, _count), _movie.get('tagline',''))
            _setProperty('%s.%d.Runtime'         % (PROPERTY, _count), runtime)
            _setProperty('%s.%d.Rating'          % (PROPERTY, _count), str(round(float(_movie.get('rating','0')),1)))
            _setProperty('%s.%d.UserRating'       % (PROPERTY, _count), str(_movie.get('userrating','0')))
            _setProperty('%s.%d.Trailer'         % (PROPERTY, _count), _movie.get('trailer',''))
            _setProperty('%s.%d.MPAA'            % (PROPERTY, _count), _movie.get('mpaa',''))
            _setProperty('%s.%d.Director'        % (PROPERTY, _count), ' / '.join(_movie.get('director','')))
            _setProperty('%s.%d.Art(thumb)'      % (PROPERTY, _count), art.get('thumb',''))
            _setProperty('%s.%d.Art(poster)'     % (PROPERTY, _count), art.get('poster',''))
            _setProperty('%s.%d.Art(fanart)'     % (PROPERTY, _count), art.get('fanart',''))
            _setProperty('%s.%d.Art(clearlogo)'  % (PROPERTY, _count), art.get('clearlogo',''))
            _setProperty('%s.%d.Art(clearart)'   % (PROPERTY, _count), art.get('clearart',''))
            _setProperty('%s.%d.Art(landscape)'  % (PROPERTY, _count), art.get('landscape',''))
            _setProperty('%s.%d.Art(banner)'     % (PROPERTY, _count), art.get('banner',''))
            _setProperty('%s.%d.Art(discart)'    % (PROPERTY, _count), art.get('discart',''))                
            _setProperty('%s.%d.Resume'          % (PROPERTY, _count), resume)
            _setProperty('%s.%d.PercentPlayed'   % (PROPERTY, _count), played)
            _setProperty('%s.%d.Watched'         % (PROPERTY, _count), watched)
            _setProperty('%s.%d.File'            % (PROPERTY, _count), _movie.get('file',''))
            _setProperty('%s.%d.Path'            % (PROPERTY, _count), path)
            _setProperty('%s.%d.Play'            % (PROPERTY, _count), play)
            _setProperty('%s.%d.VideoCodec'      % (PROPERTY, _count), streaminfo['videocodec'])
            _setProperty('%s.%d.VideoResolution' % (PROPERTY, _count), streaminfo['videoresolution'])
            _setProperty('%s.%d.VideoAspect'     % (PROPERTY, _count), streaminfo['videoaspect'])
            _setProperty('%s.%d.AudioCodec'      % (PROPERTY, _count), streaminfo['audiocodec'])
            _setProperty('%s.%d.AudioChannels'   % (PROPERTY, _count), str(streaminfo['audiochannels']))

        if _count != LIMIT:
            while _count < LIMIT:
                _count += 1
            _setProperty('%s.%d.Title'       % (PROPERTY, _count), '')
    else:
        log('## PLAYLIST {} COULD NOT BE LOADED ##'.format(PLAYLIST))
        log('JSON RESULT {}'.format(_json_pl_response))


def _getMusicVideosFromPlaylist() -> None:
    """ retrieves music video info from Kodi library and sets properties

    If a music video playlist is not provided uses music video titles node
    """
    global LIMIT
    global METHOD
    global MENU
    global PLAYLIST
    global PROPERTY
    global RESUME
    global REVERSE
    global SORTBY
    global UNWATCHED
    _result = []
    _total = 0
    _unwatched = 0
    _watched = 0
    # Request database using JSON
    if PLAYLIST == '':
        PLAYLIST = 'musicdb://musicvideos/titles'
    if JSON_RPC_NEXUS:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
                '"method": "Files.GetDirectory", '
                '"params": '
                    '{"directory": "%s", '
                    '"media": "video", '
                    '"properties": '
                        '["title", '
                        '"playcount", '
                        '"year", '
                        '"genre", '
                        '"studio", '
                        '"album", '
                        '"artist",  '
                        '"track", '
                        '"plot", '
                        '"tag", '
                        '"rating", '
                        '"userrating", '
                        '"runtime", '
                        '"file", '
                        '"lastplayed", '
                        '"resume", '
                        '"art", '
                        '"streamdetails", '
                        '"director", '
                        '"dateadded"]'
                    '}, '
                '"id": 1}' % (PLAYLIST))
    else:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
                '"method": "Files.GetDirectory", '
                '"params": '
                    '{"directory": "%s", '
                    '"media": "video", '
                    '"properties": '
                        '["title", '
                        '"playcount", '
                        '"year", '
                        '"genre", '
                        '"studio", '
                        '"album", '
                        '"artist",  '
                        '"track", '
                        '"plot", '
                        '"tag", '
                        '"rating", '
                        '"runtime", '
                        '"file", '
                        '"lastplayed", '
                        '"resume", '
                        '"art", '
                        '"streamdetails", '
                        '"director", '
                        '"dateadded"]'
                    '}, '
                '"id": 1}' % (PLAYLIST))
    _json_pl_response: dict = json.loads(_json_query)
    # If request return some results
    _files = _json_pl_response.get('result', {}).get('files')
    if _files:
        for _item in _files:
            if MONITOR.abortRequested():
                break
            _playcount = _item['playcount']
            if RESUME == 'True':
                _resume = _item['resume']['position']
            else:
                _resume = 0
            _total += 1
            if _playcount == 0:
                _unwatched += 1
            else:
                _watched += 1
            if (UNWATCHED == 'False' and RESUME == 'False') or (UNWATCHED == 'True' and _playcount == 0) or (RESUME == 'True' and _resume != 0):
                _result.append(_item)
        _setVideoProperties(_total, _watched, _unwatched)
        _count = 0
        if METHOD == 'Last':
            _result = sorted(_result, key=itemgetter(
                'dateadded'), reverse=True)
        elif METHOD == 'Playlist':
            _result = sorted(_result, key=itemgetter(SORTBY), reverse=REVERSE)
        else:
            random.shuffle(_result, random.random)
        for _musicvid in _result:
            if MONITOR.abortRequested() or _count == LIMIT:
                break
            _count += 1
            _json_query = xbmc.executeJSONRPC(
                '{"jsonrpc": "2.0", '
                    '"method": "VideoLibrary.GetMusicVideoDetails", '
                    '"params": '
                    '{"properties": ["streamdetails"], "musicvideoid":%s }, '
                    '"id": 1}' % (_musicvid['id']))
            _json_query = json.loads(_json_query)
            if 'musicvideodetails' in _json_query['result']:
                item = _json_query['result']['musicvideodetails']
                _musicvid['streamdetails'] = item['streamdetails']
            if _musicvid['resume']['position'] > 0 and float(_musicvid['resume']['total']) > 0:
                resume = 'true'
                played = '%s%%' % int(
                    (float(_musicvid['resume']['position']) / float(_musicvid['resume']['total'])) * 100)
            else:
                resume = 'false'
                played = '0%'
            if _musicvid['playcount'] >= 1:
                watched = 'true'
            else:
                watched = 'false'
            path = media_path(_musicvid['file'])
            play = 'RunScript(' + __addonid__ + \
                ',musicvideoid=' + str(_musicvid.get('id')) + ')'
            art = _musicvid['art']
            streaminfo = media_streamdetails(_musicvid['file'].lower(),
                                             _musicvid['streamdetails'])
            # Get runtime from streamdetails or from NFO
            if streaminfo['duration'] != 0:
                runtime = str(int((streaminfo['duration'] / 60) + 0.5))
                runtimesecs = (str(streaminfo['duration'] // 60) + ':'
                            + '{:02d}'.format(streaminfo['duration'] % 60))
            else:
                if isinstance(_musicvid['runtime'], int):
                    runtime = str(int((_musicvid['runtime'] / 60) + 0.5))
                    runtimesecs = (str(_musicvid['runtime'] // 60) + ':'
                                + '{:02d}'.format(_musicvid['runtime'] % 60))
                else:
                    runtime = _musicvid['runtime']
            # Set window properties
            _setProperty('%s.%d.DBID'            % (PROPERTY, _count), str(_musicvid.get('id')))
            _setProperty('%s.%d.Title'           % (PROPERTY, _count), _musicvid.get('title',''))
            _setProperty('%s.%d.Year'            % (PROPERTY, _count), str(_musicvid.get('year','')))
            _setProperty('%s.%d.Genre'           % (PROPERTY, _count), ' / '.join(_musicvid.get('genre','')))
            _setProperty('%s.%d.Studio'          % (PROPERTY, _count), ' / '.join(_musicvid.get('studio','')))
            _setProperty('%s.%d.Artist'          % (PROPERTY, _count), ' / '.join(_musicvid.get('artist','')))
            _setProperty('%s.%d.Album'           % (PROPERTY, _count), _musicvid.get('album',''))
            _setProperty('%s.%d.Track'           % (PROPERTY, _count), str(_musicvid.get('track','')))
            _setProperty('%s.%d.Rating'          % (PROPERTY, _count), str(_musicvid.get('rating','')))
            _setProperty('%s.%d.UserRating'      % (PROPERTY, _count), str(_musicvid.get('userrating','')))
            _setProperty('%s.%d.Plot'            % (PROPERTY, _count), _musicvid.get('plot',''))
            _setProperty('%s.%d.Tag'             % (PROPERTY, _count), ' / '.join(_musicvid.get('tag','')))
            _setProperty('%s.%d.Runtime'         % (PROPERTY, _count), runtime)
            _setProperty('%s.%d.Runtimesecs'     % (PROPERTY, _count), runtimesecs)
            _setProperty('%s.%d.Director'        % (PROPERTY, _count), ' / '.join(_musicvid.get('director','')))
            _setProperty('%s.%d.Art(thumb)'      % (PROPERTY, _count), art.get('thumb',''))
            _setProperty('%s.%d.Art(poster)'     % (PROPERTY, _count), art.get('poster',''))
            _setProperty('%s.%d.Art(fanart)'     % (PROPERTY, _count), art.get('fanart',''))
            _setProperty('%s.%d.Art(clearlogo)'  % (PROPERTY, _count), art.get('clearlogo',''))
            _setProperty('%s.%d.Art(clearart)'   % (PROPERTY, _count), art.get('clearart',''))
            _setProperty('%s.%d.Art(landscape)'  % (PROPERTY, _count), art.get('landscape',''))
            _setProperty('%s.%d.Art(banner)'     % (PROPERTY, _count), art.get('banner',''))
            _setProperty('%s.%d.Art(discart)'    % (PROPERTY, _count), art.get('discart',''))                
            _setProperty('%s.%d.Resume'          % (PROPERTY, _count), resume)
            _setProperty('%s.%d.PercentPlayed'   % (PROPERTY, _count), played)
            _setProperty('%s.%d.Watched'         % (PROPERTY, _count), watched)
            _setProperty('%s.%d.File'            % (PROPERTY, _count), _musicvid.get('file',''))
            _setProperty('%s.%d.Path'            % (PROPERTY, _count), path)
            _setProperty('%s.%d.Play'            % (PROPERTY, _count), play)
            _setProperty('%s.%d.VideoCodec'      % (PROPERTY, _count), streaminfo['videocodec'])
            _setProperty('%s.%d.VideoResolution' % (PROPERTY, _count), streaminfo['videoresolution'])
            _setProperty('%s.%d.VideoAspect'     % (PROPERTY, _count), streaminfo['videoaspect'])
            _setProperty('%s.%d.AudioCodec'      % (PROPERTY, _count), streaminfo['audiocodec'])
            _setProperty('%s.%d.AudioChannels'   % (PROPERTY, _count), str(streaminfo['audiochannels']))

        if _count != LIMIT:
            while _count < LIMIT:
                _count += 1
                _setProperty('%s.%d.Title' % (PROPERTY, _count), '')
    else:
        log('## PLAYLIST {} COULD NOT BE LOADED ##'.format(PLAYLIST))
        log('JSON RESULT {}'.format(_json_pl_response))


def _getEpisodesFromPlaylist() -> None:
    """retrieves episodes playlist info from Kodi library and sets properties

    """
    global LIMIT
    global METHOD
    global PLAYLIST
    global RESUME
    global REVERSE
    global SORTBY
    global UNWATCHED
    global PROPERTY
    _result = []
    _total = 0
    _unwatched = 0
    _watched = 0
    _tvshows = 0
    _tvshowid = []
    # Request database using JSON
    if JSON_RPC_NEXUS:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "Files.GetDirectory", '
            '"params": '
                '{"directory": "%s", '
                '"media": "video", '
                '"properties": '
                    '["title", '
                    '"playcount", '
                    '"season", '
                    '"episode", '
                    '"showtitle", '
                    '"plot", '
                    '"file", '
                    '"studio", '
                    '"mpaa", '
                    '"rating", '
                    '"userrating", '
                    '"resume", '
                    '"runtime", '
                    '"tvshowid", '
                    '"art", '
                    '"streamdetails", '
                    '"firstaired", '
                    '"dateadded"] '
                '}, '
            '"id": 1}' % (PLAYLIST))
    else:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "Files.GetDirectory", '
            '"params": '
                '{"directory": "%s", '
                '"media": "video", '
                '"properties": '
                    '["title", '
                    '"playcount", '
                    '"season", '
                    '"episode", '
                    '"showtitle", '
                    '"plot", '
                    '"file", '
                    '"studio", '
                    '"mpaa", '
                    '"rating", '
                    '"resume", '
                    '"runtime", '
                    '"tvshowid", '
                    '"art", '
                    '"streamdetails", '
                    '"firstaired", '
                    '"dateadded"] '
                '}, '
            '"id": 1}' % (PLAYLIST))
    _json_pl_response = json.loads(_json_query)
    _files = _json_pl_response.get('result', {}).get('files')
    if _files:
        for _file in _files:
            if MONITOR.abortRequested():
                break
            if _file['type'] == 'tvshow':
                _tvshows += 1
                # Playlist return TV Shows - Need to get episodes
                if JSON_RPC_NEXUS:
                    _json_query = xbmc.executeJSONRPC(
                        '{"jsonrpc": "2.0", '
                        '"method": "VideoLibrary.GetEpisodes", '
                        '"params": '
                            '{ "tvshowid": %s, '
                            '"properties": '
                                '["title", '
                                '"playcount", '
                                '"season", '
                                '"episode", '
                                '"showtitle", '
                                '"plot", '
                                '"file", '
                                '"rating", '
                                '"userrating", '
                                '"resume", '
                                '"runtime", '
                                '"tvshowid", '
                                '"art", '
                                '"streamdetails", '
                                '"firstaired", '
                                '"dateadded"] '
                            '}, '
                        '"id": 1}' % (_file['id']))
                else:
                    _json_query = xbmc.executeJSONRPC(
                        '{"jsonrpc": "2.0", '
                        '"method": "VideoLibrary.GetEpisodes", '
                        '"params": '
                            '{ "tvshowid": %s, '
                            '"properties": '
                                '["title", '
                                '"playcount", '
                                '"season", '
                                '"episode", '
                                '"showtitle", '
                                '"plot", '
                                '"file", '
                                '"rating", '
                                '"resume", '
                                '"runtime", '
                                '"tvshowid", '
                                '"art", '
                                '"streamdetails", '
                                '"firstaired", '
                                '"dateadded"] '
                            '}, '
                        '"id": 1}' % (_file['id']))
                _json_response = json.loads(_json_query)
                _episodes = _json_response.get('result', {}).get('episodes')
                if _episodes:
                    for _episode in _episodes:
                        if MONITOR.abortRequested():
                            break
                        # Add TV Show fanart and thumbnail for each episode
                        art = _episode['art']
                        # Add episode ID when playlist type is TVShow
                        _episode['id'] = _episode['episodeid']
                        _episode['tvshowfanart'] = art.get('tvshow.fanart')
                        _episode['tvshowthumb'] = art.get('thumb')
                        # Set MPAA and studio for all episodes
                        _episode['mpaa'] = _file['mpaa']
                        _episode['studio'] = _file['studio']
                        _total, _watched, _unwatched, _result = _watchedOrResume(
                            _total, _watched, _unwatched, _result, _episode)
                else:
                    log('## PLAYLIST {} COULD NOT BE LOADED ##'.format(PLAYLIST))
                    log('JSON RESULT {}'.format(_json_response))
            if _file['type'] == 'episode':
                _id = _file['tvshowid']
                if _id not in _tvshowid:
                    _tvshows += 1
                    _tvshowid.append(_id)
                # Playlist return TV Shows - Nothing else to do
                _total, _watched, _unwatched, _result = _watchedOrResume(
                    _total, _watched, _unwatched, _result, _file)
        _setVideoProperties(_total, _watched, _unwatched)
        _setTvShowsProperties(_tvshows)
        _count = 0
        if METHOD == 'Last':
            _result = sorted(_result, key=itemgetter(
                'dateadded'), reverse=True)
        elif METHOD == 'Playlist':
            _result = sorted(_result, key=itemgetter(SORTBY), reverse=REVERSE)
        else:
            random.shuffle(_result, random.random)
        for _episode in _result:
            if MONITOR.abortRequested() or _count == LIMIT:
                break
            _count += 1
            '''
            if _episode.get('tvshowid'):
                _json_query = xbmc.executeJSONRPC('{"jsonrpc": "2.0", "method": "VideoLibrary.GetTVShowDetails", "params": { "tvshowid": %s, "properties": ["title", "fanart", "thumbnail"] }, "id": 1}' %(_episode['tvshowid']))
                _json_pl_response = json.loads(_json_query)
                _tvshow = _json_pl_response.get('result', {}).get('tvshowdetails')
            '''
            _setEpisodeProperties(_episode, _count)
        if _count != LIMIT:
            while _count < LIMIT:
                _count += 1
                _setEpisodeProperties(None, _count)
    else:
        log('# 01 # PLAYLIST {} COULD NOT BE LOADED ##'.format(PLAYLIST))
        log('JSON RESULT {}'.format(_json_pl_response))


def _getEpisodes() -> None:
    """retrieves episode library node info from Kodi library and sets properties

    Returns:
        None
    """
    global LIMIT
    global METHOD
    global RESUME
    global REVERSE
    global SORTBY
    global UNWATCHED
    _result = []
    _total = 0
    _unwatched = 0
    _watched = 0
    _tvshows = 0
    _tvshowid = []
    # Request database using JSON
    if JSON_RPC_NEXUS:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "VideoLibrary.GetEpisodes", '
            '"params": '
                '{ "properties": '
                    '["title", '
                    '"playcount", '
                    '"season", '
                    '"episode", '
                    '"showtitle", '
                    '"plot", '
                    '"file", '
                    '"studio", '
                    '"mpaa", '
                    '"rating", '
                    '"userrating", '
                    '"resume", '
                    '"runtime", '
                    '"tvshowid", '
                    '"art", '
                    '"streamdetails", '
                    '"firstaired", '
                    '"dateadded"]'
                '}, '
            '"id": 1}')
    else:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "VideoLibrary.GetEpisodes", '
            '"params": '
                '{ "properties": '
                    '["title", '
                    '"playcount", '
                    '"season", '
                    '"episode", '
                    '"showtitle", '
                    '"plot", '
                    '"file", '
                    '"studio", '
                    '"mpaa", '
                    '"rating", '
                    '"resume", '
                    '"runtime", '
                    '"tvshowid", '
                    '"art", '
                    '"streamdetails", '
                    '"firstaired", '
                    '"dateadded"]'
                '}, '
            '"id": 1}')
    _json_pl_response = json.loads(_json_query)
    # If request return some results
    _episodes = _json_pl_response.get('result', {}).get('episodes')
    if _episodes:
        for _item in _episodes:
            if MONITOR.abortRequested():
                break
            _id = _item['tvshowid']
            if _id not in _tvshowid:
                _tvshows += 1
                _tvshowid.append(_id)
            # Add episode ID
            _item['id'] = _item['episodeid']
            _total, _watched, _unwatched, _result = _watchedOrResume(
                _total, _watched, _unwatched, _result, _item)
        _setVideoProperties(_total, _watched, _unwatched)
        _setTvShowsProperties(_tvshows)
        _count = 0
        if METHOD == 'Last':
            _result = sorted(_result, key=itemgetter(
                'dateadded'), reverse=True)
        elif METHOD == 'Playlist':
            _result = sorted(_result, key=itemgetter(SORTBY), reverse=REVERSE)
        else:
            random.shuffle(_result, random.random)
        for _episode in _result:
            if MONITOR.abortRequested() or _count == LIMIT:
                break
            _count += 1
            _setEpisodeProperties(_episode, _count)
        if _count != LIMIT:
            while _count < LIMIT:
                _count += 1
                _setEpisodeProperties(None, _count)
    else:
        log('## PLAYLIST {} COULD NOT BE LOADED ##'.format(PLAYLIST))
        log('JSON RESULT {}'.format(_json_pl_response))


def _getMusicFromPlaylist() -> None:
    """gets albums/songs from an album/songs playlist and retrieves libary data for them

    The album details are provided as window properties.  If a playlist is not 
    provided uses library songs node.  Artist and mixed playlists not supported
    """

    global LIMIT
    global METHOD
    global PLAYLIST
    global REVERSE
    global SORTBY
    _result = []
    _artists = 0
    _artistsid = []
    _albums = 0
    _albumslist = []
    _albumsid = []
    _songs = 0
    _songslist = []
    # Request database using JSON
    if PLAYLIST == '':
        PLAYLIST = 'musicdb://songs/'
    #_json_query = xbmc.executeJSONRPC('{"jsonrpc": "2.0", "method": "Files.GetDirectory", "params": {"directory": "%s", "media": "music", "properties": ["title", "description", "albumlabel", "artist", "genre", "year", "thumbnail", "fanart", "rating", "userrating", "playcount", "dateadded"]}, "id": 1}' %(PLAYLIST))
    if METHOD == 'Random':
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "Files.GetDirectory", '
            '"params": '
                '{"directory": "%s", '
                '"media": "music", '
                '"properties": ["dateadded"], '
                '"sort": {"method": "random"}}, '
            '"id": 1}' % (PLAYLIST))
    elif METHOD == 'Last':
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "Files.GetDirectory", '
            '"params": '
                '{"directory": "%s", '
                '"media": "music", '
                '"properties": ["dateadded"], '
                '"sort": '
                    '{"order": "descending", '
                    '"method": "dateadded"}}, '
            '"id": 1}' % (PLAYLIST))
    elif METHOD == 'Playlist':
        order = 'ascending'
        if REVERSE:
            order = 'descending'
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "Files.GetDirectory", '
            '"params": '
                '{"directory": "%s", '
                '"media": "music", '
                '"properties": ["dateadded"], '
                '"sort": '
                    '{"order": "%s", '
                    '"method": "%s"}}, '
            '"id": 1}' % (PLAYLIST, order, SORTBY))
    else:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "Files.GetDirectory", '
            '"params": '
                '{"directory": "%s", '
                '"media": "music", '
                '"properties": ["dateadded"]}, '
            '"id": 1}' % (PLAYLIST))
    _json_pl_response = json.loads(_json_query)
    # If request return some results
    _files: List[dict] = _json_pl_response.get('result', {}).get('files')
    #  Music type can be either album or song based on playlist type
    if _files and _files[0].get('type') == 'album':
        for _file in _files:
            if MONITOR.abortRequested():
                break
            if _file['type'] == 'album':
                _albumslist.append(_file)
                _albumid = _file['id']
                # Album playlist so get path from songs
                _json_query = xbmc.executeJSONRPC(
                    '{"jsonrpc":"2.0", '
                    '"method":"AudioLibrary.GetSongs", '
                    '"params":'
                        '{"filter":{"albumid": %s}, '
                        '"properties":["artistid"]}, ' 
                    '"id": 1}' % _albumid)
                _json_pl_response = json.loads(_json_query)
                _result = _json_pl_response.get('result', {}).get('songs')
                if _result:
                    _songs += len(_result)
                    _artistid = _result[0]['artistid']
                    if _artistid not in _artistsid:
                        _artists += 1
                        _artistsid.append(_artistid)
            #_albumid = _file.get('albumid', _file.get('id'))
            #_albumpath = os.path.split(_file['file'])[0]
            #_artistpath = os.path.split(_albumpath)[0]
            #_songs += 1
            '''
            if _albumid not in _albumsid:
                _file['id'] = _albumid
                _file['albumPath'] = _albumpath
                _file['artistPath'] = _artistpath
                _albumslist.append(_file)
                _albumsid.append(_albumid)
            '''
        _setMusicProperties(_artists, len(_files), _songs)
        if METHOD == 'Last':
            _albumslist = sorted(
                _albumslist, key=itemgetter('dateadded'), reverse=True)
        else:
            random.shuffle(_albumslist, random.random)
        _count = 0
        for _album in _albumslist:
            if MONITOR.abortRequested() or _count == LIMIT:
                break
            _count += 1
            _albumid = _album['id']
            if JSON_RPC_NEXUS:
                _json_query = xbmc.executeJSONRPC(
                    '{"jsonrpc": "2.0", '
                    '"method": "AudioLibrary.GetAlbumDetails", '
                    '"params":'
                        '{"albumid": %s, '
                        '"properties":'
                            '["title", '
                            '"description", '
                            '"albumlabel", '
                            '"theme", '
                            '"mood", '
                            '"style", '
                            '"type", '
                            '"artist", '
                            '"genre", '
                            '"year", '
                            '"thumbnail", '
                            '"fanart", '
                            '"rating", '
                            '"userrating", '
                            '"playcount"]}, '
                    '"id": 1}' % _albumid)
            else:
                _json_query = xbmc.executeJSONRPC(
                    '{"jsonrpc": "2.0", '
                    '"method": "AudioLibrary.GetAlbumDetails", '
                    '"params":'
                        '{"albumid": %s, '
                        '"properties":'
                            '["title", '
                            '"description", '
                            '"albumlabel", '
                            '"theme", '
                            '"mood", '
                            '"style", '
                            '"type", '
                            '"artist", '
                            '"genre", '
                            '"year", '
                            '"thumbnail", '
                            '"fanart", '
                            '"rating", '
                            '"playcount"]}, '
                    '"id": 1}' % _albumid)
            _json_pl_response = json.loads(_json_query)
            # If request return some results
            _album: dict = _json_pl_response.get(
                'result', {}).get('albumdetails')
            _setAlbumPROPERTIES(_album, _count)
        if _count <= LIMIT:
            while _count < LIMIT:
                _count += 1
                _setAlbumPROPERTIES(None, _count)
    elif _files and _files[0].get('type') == 'song':
        for _file in _files:
            if MONITOR.abortRequested():
                break
            _songid = _file['id']
            if JSON_RPC_NEXUS:
                _json_query = xbmc.executeJSONRPC(
                    '{"jsonrpc": "2.0", '
                    '"method": "AudioLibrary.GetSongDetails", '
                    '"params":'
                        '{"songid": %s, '
                        '"properties":'
                            '["title", '
                            '"artist", '
                            '"artistid", '
                            '"dateadded", '
                            '"genre", '
                            '"year", '
                            '"rating", '
                            '"album", '
                            '"albumid", '
                            '"track", '
                            '"duration", '
                            '"comment", '
                            '"thumbnail", '
                            '"fanart", '
                            '"userrating", '
                            '"playcount"]}, '
                    '"id": 1}' % _songid)
            else:
                _json_query = xbmc.executeJSONRPC(
                    '{"jsonrpc": "2.0", '
                    '"method": "AudioLibrary.GetSongDetails", '
                    '"params":'
                        '{"songid": %s, '
                        '"properties":'
                            '["title", '
                            '"artist", '
                            '"artistid", '
                            '"dateadded", '
                            '"genre", '
                            '"year", '
                            '"rating", '
                            '"album", '
                            '"albumid", '
                            '"track", '
                            '"duration", '
                            '"comment", '
                            '"thumbnail", '
                            '"fanart", '
                            '"playcount"]}, '
                    '"id": 1}' % _songid)
            _json_pl_response = json.loads(_json_query)
            _result: dict = _json_pl_response.get(
                'result', {}).get('songdetails')
            if _result:
                _songslist.append(_result)
                for _artistid in _result['artistid']:
                    if _artistid not in _artistsid:
                        _artists += 1
                        _artistsid.append(_artistid)
                if _result['albumid'] not in _albumslist:
                    _albums += 1
                    _albumslist.append(_result['albumid'])
        _setMusicProperties(_artists, _albums, len(_files))
        if METHOD == 'Last':
            _songslist = sorted(_songslist, key=itemgetter(
                'dateadded'), reverse=True)
        else:
            random.shuffle(_songslist, random.random)
        _count = 0
        for _song in _songslist:
            if MONITOR.abortRequested() or _count == LIMIT:
                break
            _count += 1
            _setSongPROPERTIES(_song, _count)
        if _count <= LIMIT:
            while _count < LIMIT:
                _count += 1
                _setSongPROPERTIES(None, _count)
    else:
        log('## PLAYLIST {} COULD NOT BE LOADED ##'.format(PLAYLIST))
        log('JSON RESULT {}'.format(_json_pl_response))


def _clearProperties() -> None:
    """Clears any existing window properties for the current playlist

    Returns:
        None
    """
    global WINDOW
    # Reset window Properties
    WINDOW.clearProperty('%s.Loaded' % (PROPERTY))
    WINDOW.clearProperty('%s.Count' % (PROPERTY))
    WINDOW.clearProperty('%s.Watched' % (PROPERTY))
    WINDOW.clearProperty('%s.Unwatched' % (PROPERTY))
    WINDOW.clearProperty('%s.Artists' % (PROPERTY))
    WINDOW.clearProperty('%s.Albums' % (PROPERTY))
    WINDOW.clearProperty('%s.Songs' % (PROPERTY))
    WINDOW.clearProperty('%s.Type' % (PROPERTY))


def _setMusicProperties(_artists: int, _albums: int, _songs: int) -> None:
    """sets summary porperties of albums in window properties

    Args:
        _artists (int): number of artists
        _albums (int): number of albums
        _songs (int): number of songs
    """
    global PROPERTY
    global WINDOW
    global TYPE
    # Set window Properties
    _setProperty('%s.Artists' % (PROPERTY), str(_artists))
    _setProperty('%s.Albums' % (PROPERTY), str(_albums))
    _setProperty('%s.Songs' % (PROPERTY), str(_songs))
    _setProperty('%s.Type' % (PROPERTY), TYPE)


def _setVideoProperties(_total: int, _watched: int, _unwatched: int) -> None:
    """Sets summary info of videos in window properties

    Args:
        _total (int): Total number of items
        _watched (int): Subtotal items watched
        _unwatched (int): Subtotal items unwatched
    """
    global PROPERTY
    global WINDOW
    global TYPE
    # Set window Properties
    _setProperty('%s.Count' % (PROPERTY), str(_total))
    _setProperty('%s.Watched' % (PROPERTY), str(_watched))
    _setProperty('%s.Unwatched' % (PROPERTY), str(_unwatched))
    _setProperty('%s.Type' % (PROPERTY), TYPE)


def _setTvShowsProperties(_tvshows) -> None:
    """sets tv show-level porperties

    Args:
        _tvshows(_type_): _description_
    """
    global PROPERTY
    global WINDOW
    # Set window Properties
    _setProperty('%s.TvShows' % (PROPERTY), str(_tvshows))


def _setEpisodeProperties(_episode, _count) -> None:
    """sets Kodi summary window properties for episodes

    Args:
        _episode (dict): details for an episode
        _count (_type_): episode index
    """
    if _episode:
        _json_query = xbmc.executeJSONRPC(
            '{"jsonrpc": "2.0", '
            '"method": "VideoLibrary.GetEpisodeDetails", '
            '"params": '
                '{"properties": ["streamdetails"], '
                '"episodeid":%s }, '
            '"id": 1}' % (_episode['id']))
        _json_query = json.loads(_json_query)
        if 'episodedetails' in _json_query['result']:
            item = _json_query['result']['episodedetails']
            _episode['streamdetails'] = item['streamdetails']
        episode = ('%.2d' % float(_episode['episode']))
        season = '%.2d' % float(_episode['season'])
        episodeno = 's%se%s' % (season, episode)
        rating = str(round(float(_episode['rating']), 1))
        if 'userrating' in _episode:
            userrating = str(_episode['userrating'])
        else:
            userrating = ''
        if _episode['resume']['position'] > 0 and float(_episode['resume']['total']) > 0:
            resume = 'true'
            played = '%s%%' % int(
                (float(_episode['resume']['position']) / float(_episode['resume']['total'])) * 100)
        else:
            resume = 'false'
            played = '0%'
        art = _episode['art']
        path = media_path(_episode['file'])
        play = 'RunScript(' + __addonid__ + ',episodeid=' + \
            str(_episode.get('id')) + ')'
        runtime = str(int((_episode['runtime'] / 60) + 0.5))
        streaminfo = media_streamdetails(_episode['file'].lower(),
                                         _episode['streamdetails'])
        _setProperty('%s.%d.DBID'                  % (PROPERTY, _count), str(_episode.get('id')))
        _setProperty('%s.%d.Title'                 % (PROPERTY, _count), _episode.get('title',''))
        _setProperty('%s.%d.Episode'               % (PROPERTY, _count), episode)
        _setProperty('%s.%d.EpisodeNo'             % (PROPERTY, _count), episodeno)
        _setProperty('%s.%d.Season'                % (PROPERTY, _count), season)
        _setProperty('%s.%d.Plot'                  % (PROPERTY, _count), _episode.get('plot',''))
        _setProperty('%s.%d.TVshowTitle'           % (PROPERTY, _count), _episode.get('showtitle',''))
        _setProperty('%s.%d.Rating'                % (PROPERTY, _count), rating)
        _setProperty('%s.%d.UserRating'            % (PROPERTY, _count), userrating)
        _setProperty('%s.%d.Art(thumb)'            % (PROPERTY, _count), art.get('thumb',''))
        _setProperty('%s.%d.Art(tvshow.fanart)'    % (PROPERTY, _count), art.get('tvshow.fanart',''))
        _setProperty('%s.%d.Art(tvshow.poster)'    % (PROPERTY, _count), art.get('tvshow.poster',''))
        _setProperty('%s.%d.Art(tvshow.banner)'    % (PROPERTY, _count), art.get('tvshow.banner',''))
        _setProperty('%s.%d.Art(tvshow.clearlogo)' % (PROPERTY, _count), art.get('tvshow.clearlogo',''))
        _setProperty('%s.%d.Art(tvshow.clearart)'  % (PROPERTY, _count), art.get('tvshow.clearart',''))
        _setProperty('%s.%d.Art(tvshow.landscape)' % (PROPERTY, _count), art.get('tvshow.landscape',''))
        _setProperty('%s.%d.Art(fanart)'           % (PROPERTY, _count), art.get('tvshow.fanart',''))
        _setProperty('%s.%d.Art(poster)'           % (PROPERTY, _count), art.get('tvshow.poster',''))
        _setProperty('%s.%d.Art(banner)'           % (PROPERTY, _count), art.get('tvshow.banner',''))
        _setProperty('%s.%d.Art(clearlogo)'        % (PROPERTY, _count), art.get('tvshow.clearlogo',''))
        _setProperty('%s.%d.Art(clearart)'         % (PROPERTY, _count), art.get('tvshow.clearart',''))
        _setProperty('%s.%d.Art(landscape)'        % (PROPERTY, _count), art.get('tvshow.landscape',''))
        _setProperty('%s.%d.Resume'                % (PROPERTY, _count), resume)
        _setProperty('%s.%d.Watched'               % (PROPERTY, _count), _episode.get('watched',''))
        _setProperty('%s.%d.Runtime'               % (PROPERTY, _count), runtime)
        _setProperty('%s.%d.Premiered'             % (PROPERTY, _count), _episode.get('firstaired',''))
        _setProperty('%s.%d.PercentPlayed'         % (PROPERTY, _count), played)
        _setProperty('%s.%d.File'                  % (PROPERTY, _count), _episode.get('file',''))
        _setProperty('%s.%d.MPAA'                  % (PROPERTY, _count), _episode.get('mpaa',''))
        _setProperty('%s.%d.Studio'                % (PROPERTY, _count), ' / '.join(_episode.get('studio','')))
        _setProperty('%s.%d.Path'                  % (PROPERTY, _count), path)
        _setProperty('%s.%d.Play'                  % (PROPERTY, _count), play)
        _setProperty('%s.%d.VideoCodec'            % (PROPERTY, _count), streaminfo['videocodec'])
        _setProperty('%s.%d.VideoResolution'       % (PROPERTY, _count), streaminfo['videoresolution'])
        _setProperty('%s.%d.VideoAspect'           % (PROPERTY, _count), streaminfo['videoaspect'])
        _setProperty('%s.%d.AudioCodec'            % (PROPERTY, _count), streaminfo['audiocodec'])
        _setProperty('%s.%d.AudioChannels'         % (PROPERTY, _count), str(streaminfo['audiochannels']))

    else:
         _setProperty('%s.%d.Title'               % (PROPERTY, _count), '')


def _setAlbumPROPERTIES(_album: dict, _count: int) -> None:
    """Sets the window properties for playlist albums
    """
    global PROPERTY
    if _album:
        # Set window Properties
        _rating = str(_album['rating'])
        if 'userrating' in _album:
            _userrating = str(_album['userrating'])
        else:
            _userrating = ''
        if _rating == '48':
            _rating = ''
        play = 'RunScript(' + __addonid__ + ',albumid=' + \
            str(_album.get('albumid')) + ')'
        path = 'musicdb://albums/' + str(_album.get('albumid')) + '/'
        _setProperty('%s.%d.Title'       % (PROPERTY, _count), _album.get('title',''))
        _setProperty('%s.%d.Artist'      % (PROPERTY, _count), ' / '.join(_album.get('artist','')))
        _setProperty('%s.%d.Genre'       % (PROPERTY, _count), ' / '.join(_album.get('genre','')))
        _setProperty('%s.%d.Theme'       % (PROPERTY, _count), ' / '.join(_album.get('theme','')))
        _setProperty('%s.%d.Mood'        % (PROPERTY, _count), ' / '.join(_album.get('mood','')))
        _setProperty('%s.%d.Style'       % (PROPERTY, _count), ' / '.join(_album.get('style','')))
        _setProperty('%s.%d.Type'        % (PROPERTY, _count), _album.get('type',''))
        _setProperty('%s.%d.Year'        % (PROPERTY, _count), str(_album.get('year','')))
        _setProperty('%s.%d.RecordLabel' % (PROPERTY, _count), _album.get('albumlabel',''))
        _setProperty('%s.%d.Description' % (PROPERTY, _count), _album.get('description',''))
        _setProperty('%s.%d.Rating'      % (PROPERTY, _count), _rating)
        _setProperty('%s.%d.UserRating'  % (PROPERTY, _count), _userrating)
        _setProperty('%s.%d.Art(thumb)'  % (PROPERTY, _count), _album.get('thumbnail',''))
        _setProperty('%s.%d.Art(fanart)' % (PROPERTY, _count), _album.get('fanart',''))
        _setProperty('%s.%d.Play'        % (PROPERTY, _count), play)
        _setProperty('%s.%d.LibraryPath' % (PROPERTY, _count), path)
    else:
        _setProperty('%s.%d.Title'       % (PROPERTY, _count), '')


def _setSongPROPERTIES(_song: dict, _count: int) -> None:
    """Sets the window properties for playlist songs
    """
    global PROPERTY
    if _song:
        # Set window Properties
        _rating = str(_song['rating'])
        if 'userrating' in _song:
            _userrating = str(_song['userrating'])
        else:
            _userrating = ''
        if _rating == '48':
            _rating = ''
        play = 'RunScript(' + __addonid__ + ',songid=' + \
            str(_song.get('songid')) + ')'
        path = 'musicdb://songs/' + str(_song.get('songid')) + '/'
        _setProperty('%s.%d.Title'       % (PROPERTY, _count), _song.get('title',''))
        _setProperty('%s.%d.Artist'      % (PROPERTY, _count), ' / '.join(_song.get('artist','')))
        _setProperty('%s.%d.Genre'       % (PROPERTY, _count), ' / '.join(_song.get('genre','')))
        _setProperty('%s.%d.Year'        % (PROPERTY, _count), str(_song.get('year','')))
        _setProperty('%s.%d.Description' % (PROPERTY, _count), _song.get('comment',''))
        _setProperty('%s.%d.Rating'      % (PROPERTY, _count), _rating)
        _setProperty('%s.%d.UserRating'  % (PROPERTY, _count), _userrating)
        _setProperty('%s.%d.Art(thumb)'  % (PROPERTY, _count), _song.get('thumbnail',''))
        _setProperty('%s.%d.Art(fanart)' % (PROPERTY, _count), _song.get('fanart',''))
        _setProperty('%s.%d.Play'        % (PROPERTY, _count), play)
        _setProperty('%s.%d.LibraryPath' % (PROPERTY, _count), path)
    else:
        _setProperty('%s.%d.Title'       % (PROPERTY, _count), '')


def _setProperty(_property: str, _value: str) -> None:
    """Calls kodi setProperty method

    Args:
        _property (str): property key
        _value (str): value
    """
    global WINDOW
    # Set window Properties
    WINDOW.setProperty(_property, _value)


def _parse_argv() -> None:
    """Gets arguments pass by skin call to RunScript()

        Arguments are retrieved into a dict for processing.
        -  If a library item id is passed, starts playback of item 
        -  Otherwise will set script global variables
        -  If passed playlist will determine type and order
    """
    try:
        params = dict(arg.split('=') for arg in sys.argv[1].split('&'))
    except:
        params = {}
    if params.get('movieid'):
        #xbmc.executeJSONRPC('{ "jsonrpc": "2.0", "method": "Player.Open", "params": { "item": { "movieid": %d }, "options":{ "resume": true } }, "id": 1 }' % int(params.get("movieid")))
        xbmc.executeJSONRPC('{ "jsonrpc": "2.0", '
        '"method": "Player.Open", '
        '"params": '
            '{ "item": { "movieid": %d }, '
            '"options":{ "resume": %s } }, '
        '"id": 1 }' % (
            int(params.get("movieid", "")), params.get("resume", "true")))
    elif params.get('episodeid'):
        #xbmc.executeJSONRPC('{ "jsonrpc": "2.0", "method": "Player.Open", "params": { "item": { "episodeid": %d }, "options":{ "resume": true }  }, "id": 1 }' % int(params.get("episodeid")))
        xbmc.executeJSONRPC('{ "jsonrpc": "2.0", '
                            '"method": "Player.Open", '
                            '"params": '
                                '{ "item": { "episodeid": %d }, '
                            '"options":{ "resume": %s }  }, '
                            '"id": 1 }' % (int(params.get("episodeid", "")), params.get("resume", "true")))
    elif params.get('musicvideoid'):
        xbmc.executeJSONRPC('{ "jsonrpc": "2.0", '
                            '"method": "Player.Open", '
                            '"params": { "item": { "musicvideoid": %d } }, '
                            '"id": 1 }' % int(params.get("musicvideoid")))
    elif params.get('albumid'):
        xbmc.executeJSONRPC('{ "jsonrpc": "2.0", '
                            '"method": "Player.Open", '
                            '"params": { "item": { "albumid": %d } }, '
                            '"id": 1 }' % int(params.get("albumid")))
    elif params.get('songid'):
        xbmc.executeJSONRPC(
            '{ "jsonrpc": "2.0", '
            '"method": "Player.Open", '
            '"params": { "item": { "songid": %d } }, '
            '"id": 1 }' % int(params.get("songid")))
    else:
        global METHOD
        global MENU
        global LIMIT
        global PLAYLIST
        global PROPERTY
        global RESUME
        global TYPE
        global UNWATCHED
        # Extract parameters
        for arg in sys.argv:
            param = str(arg)
            if 'limit=' in param:
                LIMIT = int(param.replace('limit=', ''))
            elif 'menu=' in param:
                MENU = param.replace('menu=', '')
            elif 'method=' in param:
                METHOD = param.replace('method=', '')
            elif 'playlist=' in param:
                PLAYLIST = param.replace('playlist=', '')
                PLAYLIST = PLAYLIST.replace('"', '')
            elif 'property=' in param:
                PROPERTY = param.replace('property=', '')
            elif 'type=' in param:
                TYPE = param.replace('type=', '')
            elif 'unwatched=' in param:
                UNWATCHED = param.replace('unwatched=', '')
                if UNWATCHED == '':
                    UNWATCHED = 'False'
            elif 'resume=' in param:
                RESUME = param.replace('resume=', '')
        if PLAYLIST != '' and xbmcvfs.exists(xbmcvfs.translatePath(PLAYLIST)):
            _getPlaylistType()
        if PROPERTY == '':
            PROPERTY = 'Playlist%s%s%s' % (METHOD, TYPE, MENU)


def media_streamdetails(filename: str, streamdetails: dict) -> dict:
    """gets stream details from Kodi library computes the video resolution

    Args:
        filename (str): filename
        streamdetails (dict): A dict from the json rpc call

    Returns:
        dict: stream details
    """
    info = {}
    video = streamdetails['video']
    audio = streamdetails['audio']
    if '3d' in filename:
        info['videoresolution'] = '3d'
    elif video:
        videowidth = video[0]['width']
        videoheight = video[0]['height']
        if (video[0]['width'] <= 720 and video[0]['height'] <= 480):
            info['videoresolution'] = '480'
        elif (video[0]['width'] <= 768 and video[0]['height'] <= 576):
            info['videoresolution'] = '576'
        elif (video[0]['width'] <= 960 and video[0]['height'] <= 544):
            info['videoresolution'] = '540'
        elif (video[0]['width'] <= 1280 and video[0]['height'] <= 720):
            info['videoresolution'] = '720'
        elif (video[0]['width'] >= 1281 and video[0]['height'] >= 721):
            info['videoresolution'] = '1080'
        else:
            info['videoresolution'] = ''
    elif (('dvd') in filename and not ('hddvd' or 'hd-dvd') in filename) or (filename.endswith('.vob' or '.ifo')):
        info['videoresolution'] = '576'
    elif (('bluray' or 'blu-ray' or 'brrip' or 'bdrip' or 'hddvd' or 'hd-dvd') in filename):
        info['videoresolution'] = '1080'
    else:
        info['videoresolution'] = '1080'
    if video and 'duration' in video[0]:
        info['duration'] = video[0]['duration']
    else:
        info['duration'] = 0
    if video:
        info['videocodec'] = video[0]['codec']
        if (video[0]['aspect'] < 1.4859):
            info['videoaspect'] = '1.33'
        elif (video[0]['aspect'] < 1.7190):
            info['videoaspect'] = '1.66'
        elif (video[0]['aspect'] < 1.8147):
            info['videoaspect'] = '1.78'
        elif (video[0]['aspect'] < 2.0174):
            info['videoaspect'] = '1.85'
        elif (video[0]['aspect'] < 2.2738):
            info['videoaspect'] = '2.20'
        else:
            info['videoaspect'] = '2.35'
    else:
        info['videocodec'] = ''
        info['videoaspect'] = ''
    if audio:
        info['audiocodec'] = audio[0]['codec']
        info['audiochannels'] = audio[0]['channels']
    else:
        info['audiocodec'] = ''
        info['audiochannels'] = ''
    return info


def media_path(path) -> str:
    """gets actual path for the media item

    Args:
        path (str): The Kodi path for the media item

    Returns:
        str: The path fixed for stacked or RARed items
    """
    # Check for stacked movies
    try:
        path: tuple = os.path.split(path)[0].rsplit(' , ', 1)[1].replace(',,', ',')
    except:
        path = os.path.split(path)[0]
    # Fixes problems with rared movies and multipath
    if path.startswith('rar://'):
        path = [os.path.split(urllib.request.url2pathname(
            path.replace('rar://', '')))[0]]
    elif path.startswith('multipath://'):
        temp_path = path.replace('multipath://', '').split('%2f/')
        path = []
        for item in temp_path:
            path.append(urllib.request.url2pathname(item))
    else:
        path = [path]
    return path[0]


# Parse argv for any preferences
_parse_argv()
# Clear Properties for playlist PROPERTY from _parse_argv()
_clearProperties()
# Get movies and fill Properties
if TYPE == 'Movie':
    _getMovies()
elif TYPE == 'Episode':
    if PLAYLIST == '':
        _getEpisodes()
    else:
        _getEpisodesFromPlaylist()
elif TYPE == 'Music':
    _getMusicFromPlaylist()
elif TYPE == 'MusicVideo':
    _getMusicVideosFromPlaylist()
if TYPE != 'Invalid':
    # skin can check this to verify properties available
    WINDOW.setProperty('{}.Loaded'.format(PROPERTY), 'true')
    log('Loading Playlist{0}{1}{2} started at {3} and took {4} (Nexus {5})'.format(METHOD, TYPE, MENU, time.strftime(
        '%Y-%m-%d %H:%M:%S', time.localtime(START_TIME)), _timeTook(START_TIME), JSON_RPC_NEXUS))
else:
    log('Unable to process the {}{} playlist'.format(METHOD, MENU))
