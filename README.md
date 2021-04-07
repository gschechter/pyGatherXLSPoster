> :warning: **I am a total noob python developer** : I am very open to feedback.  My primary goal was to write this as easy to read, over n-th performance gains so if I needed to make modifications in 6 months or 2 years it would be easier.

# GatherTown Poster Updater From XLSX (PYTHON)

Easily bulk update posters in GatherTown from XLSX file.

## How It Works

1. Python Script - Requires
    * OPENPYXL
    * REQUSTS
    * JSON
    * COLORAMA
    * CONFIGPARSER
    * OS PATH
2. Creates a basic config file if missing.  
3. Won't run if configured XLSX is missing.  XLSX File has 3 columns for `mapid`/`room name`, `caption`/`blurb`, and `url`
<img src='excel.PNG'/>
> :warning: **Make sure to have CORS on the hosting server set to allow the origin** [GatherTown Mentions this in thier API documentation](https://www.notion.so/EXTERNAL-Gather-http-API-3bbf6c59325f40aca7ef5ce14c677444#fbba4d665f814186a6c82523203b91af).

4. Let's you override the apikey/spaceid from the config.
5. Rebuilds XLSX file into something that can that be parsed easier and more efficiently to consolidate the API calls.
    ```json[{
    'mapid': 'ExampleRoom1',
    'posters': [{
        'blurb': 'poster1',
        'url': 'picture3.jpg'
    }, {
        'blurb': 'poster2',
        'url': 'picture4.jpg'
    }, {
        'blurb': 'poster3',
        'url': 'picture5.jpg'
    }]}, {
    'mapid': 'ExampleRoom2',
    'posters': [{
        'blurb': 'poster1',
        'url': 'picture6.jpg'
    }, {
        'blurb': 'poster2',
        'url': 'picture7.jpg'
    }]}]
6. Get's the map data from gatherTown's getMap endpoint, and searches on the `blurb`/`caption` for the posters it update
   * Set under the showmore toggle on the map maker.
<img src="gathertownshowmore.PNG"/>
   <img src="postercaptionblurb.PNG" />
7. Updates the poster `image` and `preview` URLs with the one from the excel file.

8. Sends the updated map data with gatherTown's setMap endpoint.

### I used `pyinstaller` to create a distributable exe for my colleagues to not have to load and configure a python runtime environment. ###
    pyinstaller -f pyGatherXLSPoster.py
