# -*- coding: utf-8 -*-
import urllib
import time
import re
import os
import datetime
from threading import Thread
import random
import json
import inspect
import urllib
import time
import decimal
import pickle
import pywinauto
import win32com.client
from flask import Flask, render_template, request, g, session, flash, \
    redirect, url_for, abort, jsonify, make_response
from openid.extensions import pape
from flask.ext.script import Manager, Command
import pythoncom


shell = win32com.client.Dispatch("WScript.Shell")





class fooBar():

    def change_trackanddelete_currentone(self):
        pwa_app = pywinauto.Application()
        w_handle = pywinauto.findwindows.find_windows(class_name='{97E27FAA-C0B3-4b8e-A693-ED7881E99FC1}')[0]
        window = pwa_app.window_(handle=w_handle)
        window.SetFocus()
        time.sleep(2)

        window.TypeKeys('^d')
        self.next()

        time.sleep(0.5)
        shell.SendKeys("{DELETE}") # Delete selected text?  Depends on context. :P
        time.sleep(0.5)
        shell.SendKeys("{ENTER}") # Delete selected text?  Depends on context. :P


    def change_trackandadd_tomusicdirectorythen_delete(self):
        pwa_app = pywinauto.Application()
        w_handle = pywinauto.findwindows.find_windows(class_name='{97E27FAA-C0B3-4b8e-A693-ED7881E99FC1}')[0]
        window = pwa_app.window_(handle=w_handle)
        window.Click()
        window.SetFocus()
        time.sleep(1)
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










#heroku scheduler commands
#python example.py fetch_eve_online_deposits

# setup flask
app = Flask(__name__)
manager = Manager(app)



if os.environ.get('SECRET_KEY') is None:
    SECRET_KEY = "ldkjflskdjf"
else:
    SECRET_KEY = str(os.environ['SECRET_KEY'])

if os.environ.get('SECRET_KEY') is None:
    DEBUG_var = True
else:  # on heroku we set debug to false
    DEBUG_var = True

app.config.update(
    SECRET_KEY=SECRET_KEY,
    DEBUG=DEBUG_var
)

#jinja
app.jinja_env.add_extension('jinja2.ext.loopcontrols')



@app.route('/')
def index():
    return render_template('index.html')

@app.route('/jumpforward')
def jumpforward():
    global playback
    pythoncom.CoInitialize()

    ProgID = "Foobar2000.Application.0.7"
    foobar_COM_object = win32com.client.Dispatch(ProgID)
    playback = foobar_COM_object.Playback
    f = fooBar()

    f.jumpForward10Seconds()
    return redirect(url_for('index'))

@app.route('/jumpbackward')
def jumpbackward():
    global playback
    pythoncom.CoInitialize()

    ProgID = "Foobar2000.Application.0.7"
    foobar_COM_object = win32com.client.Dispatch(ProgID)
    playback = foobar_COM_object.Playback
    f = fooBar()

    f.jumpBackward10Seconds()
    return redirect(url_for('index'))

@app.route('/change_trackanddelete_currentone')
def changetrackanddeletecurrentone():
    global playback
    pythoncom.CoInitialize()

    ProgID = "Foobar2000.Application.0.7"
    foobar_COM_object = win32com.client.Dispatch(ProgID)
    playback = foobar_COM_object.Playback
    f = fooBar()

    f.change_trackanddelete_currentone()
    return redirect(url_for('index'))


@app.route('/change_trackandadd_tomusicdirectorythen_delete')
def change_trackandadd_tomusicdirectorythen_delete():
    global playback
    pythoncom.CoInitialize()

    ProgID = "Foobar2000.Application.0.7"
    foobar_COM_object = win32com.client.Dispatch(ProgID)
    playback = foobar_COM_object.Playback
    f = fooBar()

    f.change_trackandadd_tomusicdirectorythen_delete()
    return redirect(url_for('index'))

@app.route('/playnext')
def playnext():
    global playback
    pythoncom.CoInitialize()

    ProgID = "Foobar2000.Application.0.7"
    foobar_COM_object = win32com.client.Dispatch(ProgID)
    playback = foobar_COM_object.Playback
    f = fooBar()

    f.next()
    return redirect(url_for('index'))

##END ROUTING


if __name__ == '__main__':
    app.run(host='0.0.0.0')