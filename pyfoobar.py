__module_name__ = "fb2kcomtest" 
__module_version__ = "1.0" 
__module_description__ = "Controls foobar2000 audio player through COM"
__module_author__ = "Holger stenger <holger.stenger@gmx.de>"
import win32gui
import time
import pywinauto
import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")

spec = "[%artist% - ]$if2(%title%,%_filename_ext%)' ['$if(%album%,%album%[ '('%date%')'],$if(%date%,%date%[ / %venue%]))[ #[%disc%/]$num(%tracknumber%,2)]']'[ '['%_length%', '%__bitrate%kbps $codec()']']"

testArtist = "Skillet"
#
# def test():
#     try:
#         ProgID = "Foobar2000.Application.0.7"
#         fb2k = win32com.client.Dispatch(ProgID)
#         print "- Testing playback interface..."
#         try:
#             playback = fb2k.Playback
#             print "-- Testing playback orders interface..."
#             try:
#                 pos = playback.Settings.PlaybackOrders
#                 print "--- Enumerating available playback orders:"
#                 for po in pos:
#                     print po
#                 print "--- Converting name to GUID:"
#                 if pos.Count > 0:
#                     name = pos[0]
#                     guid = "<not retrieved>"
#                     try:
#                         guid = pos.GuidFromName(name)
#                     except:
#                         pass
#                     print name, "->", guid
#             except:
#                 print "There was an error."
#         except:
#             print "There was an error."
#         print "- Testing media library interface..."
#         try:
#             ml = fb2k.MediaLibrary
#             print "--- Rescanning media library."
#             ml.Rescan()
#             print "--- Getting all tracks:"
#             tl = ml.GetTracks()
#             print "There are", tl.Count, "tracks in the media library."
#             print "--- Searching tracks:"
#             tl2 = ml.GetTracks("artist IS " + testArtist)
#             print "There are", tl2.Count, "tracks by " + testArtist + " in the media library."
#             print "-- Testing track list interface..."
#             try:
#                 tl3 = tl.GetTracks("artist IS " + testArtist)
#                 print "There are", tl3.Count, "tracks by " + testArtist + " in the track list retrieved from the media library."
#             except:
#                 print "There was an error."
#         except:
#             print "There was an error."
#     except:
#         pass
#
# #test()


pwa_app = pywinauto.Application()
w_handle = pywinauto.findwindows.find_windows(class_name='{97E27FAA-C0B3-4b8e-A693-ED7881E99FC1}')[0]
window = pwa_app.window_(handle=w_handle)

ProgID = "Foobar2000.Application.0.7"
foobar_COM_object = win32com.client.Dispatch(ProgID)
playback = foobar_COM_object.Playback

class fooBar():

    def change_trackanddelete_currentone(self):
        window.Click()
        window.SetFocus()
        window.TypeKeys('^d')
        self.next()

        time.sleep(0.5)
        shell.SendKeys("{DELETE}") # Delete selected text?  Depends on context. :P
        time.sleep(0.5)
        shell.SendKeys("{ENTER}") # Delete selected text?  Depends on context. :P


    def change_trackandadd_tomusicdirectorythen_delete(self):
        window.Click()
        window.SetFocus()
        window.TypeKeys('^d')
        self.next()

        time.sleep(0.5)
        shell.SendKeys('%{f}')
        time.sleep(0.1)
        shell.SendKeys('n')
        time.sleep(0.1)
        shell.SendKeys('n')
        time.sleep(0.1)
        shell.SendKeys('n')
        time.sleep(0.1)
        shell.SendKeys('n')
        time.sleep(0.1)
        shell.SendKeys("{ENTER}") # Delete selected text?  Depends on context. :P
        time.sleep(0.5)
        shell.SendKeys("m") # Delete selected text?  Depends on context. :P
        time.sleep(0.5)
        shell.SendKeys("{DELETE}") # Delete selected text?  Depends on context. :P
        time.sleep(0.5)
        shell.SendKeys("{ENTER}") # Delete selected text?  Depends on context. :P


    def isPlaying(self):
        return playback.IsPlaying

    def seekToPosition(self,seconds):
        if self.isPlaying():
            playback.Seek(seconds)

    def jumpForward10Seconds(self):
        if self.isPlaying():
            playback.Seek(playback.Position+10)

    def jumpBackward10Seconds(self):
        if self.isPlaying():
            playback.Seek(playback.Position-10)

    def play(self):
        playback.Start()

    def stop(self):
           playback.Stop()

    def pauseplay(self):
           playback.Pause()

    def isPaused(self):
          return playback.IsPaused

    def next(self):
           playback.Next()

    def previous(self):
           playback.Previous()

    def playRandom(self):
           playback.Random()

    def seekPosition(self):
           return playback.Position

    def lengthOfTrack(self):
            return str(playback.FormatTitle("[%length%]"))


    def currentVolumeLevel(self):
            return str(playback.Settings.Volume) + "dB"

    def mute(self):
            playback.Settings.Volume = -100

    def setVolumeLevel(self,value):
            '''Set volume level to given value
               0dB corresponds to MAX_VALUE and -100dB corresponds to MIN_VALUE
               So, -100 <= value <= 0'''
            playback.Settings.Volume = value

    def currentActivePlaylist(self):
            return str(foobar_COM_object.Playlists.ActivePlaylist.Name)

    def getCurrentTrack(self):
            if self.isPlaying():
                    track = str(playback.FormatTitle("[%title%]"))
                    if len(track) == 0:
                            return "check metadata"
                    else:
                            return track
            else:
                    return "Check foobar running or not"


    def getCurrentArtist(self):
            if self.isPlaying():
                    artist = str(playback.FormatTitle("[%artist%]"))
                    if len(artist) == 0:
                            return "check metadata"
                    else:
                            return artist
            else:
                    return "Check foobar running or not"


    def getCurrentAlbum(self):
            if self.isPlaying():
                    album = str(playback.FormatTitle("[%album%]"))
                    if len(album) == 0:
                            return "check metadata"
                    else:
                            return album
            else:
                    return "Check foobar running or not"


    def isCurrentlyPlaying(self):
            return 'Currently playing "{0}" by "{1}"'.format(self.getCurrentTrack(),self.getCurrentArtist())

f = fooBar()

print f.getCurrentTrack()

print f.seekPosition()








