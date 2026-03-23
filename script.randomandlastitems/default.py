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

if __name__ == "__main__":
    from resources.lib import randomandlastitems as rali
    rali.run()
