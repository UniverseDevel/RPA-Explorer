# RPA Explorer

Graphical explorer for RenPy Archives. This tool brings ability to extract, create new or change existing RPA archives all in one window. It also provides content preview for most common files in these packages. Initial parser code was inspired by [RPATools](https://github.com/Shizmob/rpatool), so in case you find this tool usefull, go give them a thumbs up as well.

Note: This is a fan made application and there is no guarantee of further development or fixes. For video support LibVLC library is used and this library has ~300MiB in size so this is the reason why this application is so big, I haven't found a better way around this yet.

Supported file types for preview:

- Text: py, rpy~, rpy, txt, log, nfo, htm, html, xml, json, yaml, csv
- Video: 3gp, flv, mov, mp4, ogv, swf, mpg, mpeg, avi, mkv, wmv, .webm
- Audio: aac, ac3, flac, mp3, wma, wav, ogg, cpc
- Images: jpeg, jpg, bmp, tiff, png, webp, exif, ico, gif

Download link: [RPA Explorer.7z](https://github.com/UniverseDevel/RPA-Explorer/blob/master/RPA%20Explorer/bin/Release/net461/RPA%20Explorer.7z)

TODO List: [TODO.md](https://github.com/UniverseDevel/RPA-Explorer/blob/master/TODO.md)

Known Issues:

- When selecting/unselecting objects too fast will not update selections for child or parent objects, this seems to be a TreeView bug/shortcomming and there is not much I can do with it.
- Some video/audio formats will not update time played or total video time and this seems to be LibVLC library issue.

Images preview:
![1](https://user-images.githubusercontent.com/47400898/154856556-1da3d011-5631-4100-972c-f6e844967242.png)
Video preview:
![2](https://user-images.githubusercontent.com/47400898/154856560-71837ed7-899c-43bb-ab0d-3a10dd7844e8.png)
Text files preview:
![3](https://user-images.githubusercontent.com/47400898/154856564-1a588bdd-3412-491d-a070-078e17c42d19.png)

The software is provided "as is", without a warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose and non-infringement. In no event shall the authors or copyright holders be liable for any claim, damages or other liability, whether in an action of contract, tort or otherwise, arising from, out of or in connection with the software or the use or other dealings in the software.
