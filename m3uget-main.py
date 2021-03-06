# IF programme

import requests
import fileinput
import sys
import operator
from operator import itemgetter
import urllib
import zipfile
import io

# Configure your Files
fpath = 'U:\\Temp\\'
ktemp1 = 'ktemp1.tmp'
ktemp2 = 'ktemp2.tmp'
tvlist = 'tvlists.txt'

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
    sender = sender.rstrip()
    for kodi in fileinput.input(fpath + ktemp1, openhook=fileinput.hook_encoded('utf-8')):
        if sender in kodi:
            kodi = kodi.replace(sender, sender.lower())
            kodi = kodi.replace('#EXTINF', 'P#EXTINF')
            kodi = kodi.rstrip()
            fl.write(kodi)
        elif 'http' in kodi:
            fl.write(kodi)
        else:
            pass
pass
fl.close()

fl = open(fpath + ktemp1, 'w', encoding='utf-8')
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
for line in fileinput.input(fpath + ktemp1, openhook=fileinput.hook_encoded("utf-8")):
    sp = line.split(' ')
    klist2.append(sp)

slist = sorted(klist2, key=itemgetter(4))

# Step four : Make the final m3u List
fl = open(fpath + 'kodistr-f.m3u', 'a', encoding='utf-8')
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
