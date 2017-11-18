# IF programme

import requests
import fileinput
from operator import itemgetter
import zipfile
import io
import re

# Configure your Files
fpath = 'U:\\Temp\\'
ktemp1 = 'ktemp1.tmp'  # original List from internet
ktemp2 = 'ktemp2.tmp'  # File with nested loop to get all programs
ktemp3 = 'ktemp3.tmp'  # Final list just http and extif in one line
tvlist = 'tvlist.txt'  # List of all TV Stations wanted

# Get the latest m3u from IPTVSaga
url = 'http://www.iptvsaga.com/play/tv2/index.php'
# You need to get it as a kodi agent otherwise fail
headers = {'User-agent': 'Kodi/14.0 (Macintosh; Intel Mac OS X 10_10_3) App_Bitness/64 Version/14.0-Git:2014-12-23-ad747d9-dirty'}

#  I import everything into the file as binary due to the incompatibility of diverse utf-8 ö,ä and everything else
req = requests.get(url,  headers=headers)
resp = req.content
fl = open(fpath + ktemp1, 'wb')
fl.write(resp)
fl.close()

# Step one : Run List of TV Stations as loop against file and mark all stations in file

fl = open(fpath + ktemp2, 'w', encoding='utf-8')      # Need fileinput.input not open
for sender in open(fpath + tvlist, encoding='utf-8'):
    # Get the station name
    station = re.match('(.*?)XOX', sender).group(0)
    station = station[:-4]
    # Get the tvg-id
    tvgid = re.search('XOX_(.*?)_XOX', sender).group(0)
    tvgid = 'tvg-id="' + tvgid[4:-4] + '"'
    print(tvgid)
    # Now loop these two variables into the main file
    for kodi in fileinput.input(fpath + ktemp1, openhook=fileinput.hook_encoded('utf-8')):
        if station in kodi:
            kodi = kodi.replace(station, station.lower())
            kodi = kodi.replace('#EXTINF', 'P#EXTINF')
            kodi = kodi.rstrip()
            kodi = re.sub('tvg-id="(.*?)"', tvgid, kodi)
            print(kodi)
            fl.write(kodi)
        elif 'http' in kodi:
            fl.write(kodi)
        else:
            pass
pass
fl.close()

# Step two : strip list down to the real winners

fl = open(fpath + ktemp3, 'w', encoding='utf-8')
for clean in open(fpath + ktemp2, 'r', encoding='utf-8').readlines():
    if 'P#EXTINF'in clean:
        clean=clean.replace('P#EXTINF', '#EXTINF')
        fl.write(clean)
    else:
        pass
pass
fl.close()

# Step three : Make a List with each line as a further list to sort the list of lists by element 2
klist2 =[]
for line in fileinput.input(fpath + ktemp3, openhook=fileinput.hook_encoded("utf-8")):
    sp = line.split(' ')
    klist2.append(sp)

slist = sorted(klist2, key=itemgetter(4))

# Step four : Make the final m3u List
fl = open(fpath + 'kodistr-f.m3u', 'w', encoding='utf-8')
i = len(slist) - 1
while i >= 0:
    j = slist[i]
    k = ' '.join(j)
    l = k.replace('[/COLOR]', '[/COLOR]\n')
    fl.write(l)
    i = i - 1
fl.close()

# Step four : Get the IPTV Download for missing Channels
url = 'https://www.iptv4sat.com/download-iptv-germany/'
req = requests.get(url)
f = req.text
i = f.find('<a href="https://www.iptv4sat.com/download-attachment')
j = int(i+9)
k = int(i+98)
m3uf = f[j:k]
m3uf = str(m3uf)

r = requests.get(m3uf)
z = zipfile.ZipFile(io.BytesIO(r.content))
zi = str(z.infolist())
zl = zi.split("'")
zl = zl[1]
z.extractall(fpath)


# Step to do: Replace the tvgid of free4fisher with my own
str1 = 'P#EXTINF:-1 tvg-id="3plus.ch" group-title="Deutsch" tvg-logo="0135.png",[COLOR orangered]3+ hd[/COLOR]http://wilmaa.customers.cdn.iptv.ch/1/1013/index.m3u8?token=34ec24a9f1d948566691d6ee41e36f44&expires=1511006957&c=t3'
tvlist = '3+ HD'
sublist1 = '3+'
resub = 'tvg-id="' + sublist1 +'"'

str2 = re.sub('tvg-id="[0-9a-zA-Z]*.[0-9a-zA-Z]*"', resub, str1, re.I)
str3 = re.search(r'\](.*?)\[', str1)
