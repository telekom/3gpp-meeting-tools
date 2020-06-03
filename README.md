# 3gpp-meeting-tools
A set of tools to ease the execution of 3GPP meetings

The purpose of this project is to facilitate the actual
work of 3GPP SA2 delegates by automating certain
repetitive tasks caused by the meeting procedures. This includes:

* Searching of TDocs in the folder structure combining web locations, 
a local repository and a local inbox folder where contributions are 
uploaded on-the-fly
* Searching of the last agenda file in the meeting folder structure 
(both local and remote)
* Caching TDocs for offline use/review in an ordered local file structure
* Merging of TDoc comments from different colleagues to form a 
consolidated list that can be used by the delegate during the meeting
* Automaticly sorting of eMeeting and email approval emails by Agenda Item

# Table of contents
1. [Introduction](#introduction)
2. [Installation](#installation)
  1. [Installing Python](#installing-python)
  2. [Installing dependencies](#installing-dependencies)
  3. [Installing the tricky libraries before running pip](#installing-the-tricky-libraries-before-running-pip)
  4. [Configuring your firewall](#configuring-your-firewall)
3. [Running](#running)

# Introduction

Small Python application designed to make life easier for 3GPP SA2 delegates.



For detailed information, please check the [project Wiki](https://gitlab.com/ikuno/3gpp-sa2-meeting-helper/wikis/home)

# Installation

The program is written in [Python](python.org). It should work on Python 3.7 and higher. The basic requirement is that you have Python installed and you then install the application dependencies.

The application should work both in the 3GPP WLAN, Internet and corporate network (as long as you know the HTTP proxy configuration).

## Installing Python

Go to [python.org](python.org) and download the latest version of Python (3.8 as of the time this document was written).

I would recommend that you select "add Python to PATH" so that you call python by simply using the ``python`` command on the console, but that is up to you.

## Installing dependencies

While techncially, if you have the proper environment, you could download all of them by going to the folder where you saved the application and running ``pip install -r requirements.txt``, in practice some libraries may be a bit tricky to install if you do not have a developer environment (which I assume you don't).

Also, since the application has code to generate Word/Excel/Outlook calls via [COM](https://en.wikipedia.org/wiki/Component_Object_Model), you will need [pywin32](https://github.com/mhammond/pywin32/releases) for your Python release (e.g. Python 3.8 32-bit).

## Installing the tricky libraries before running pip

If the installation of some libraries failed, you most probably need to install pre-compiled libraries for some of the dependencies. Luckily, you can find all of them [here](https://www.lfd.uci.edu/~gohlke/pythonlibs/).
* [numpy](https://www.lfd.uci.edu/~gohlke/pythonlibs/#numpy)
* [pandas](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pandas)
* [lxml](https://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml)

For each of the libraries, download the correct [Wheel file](https://pythonwheels.com/). e.g. for numpy, if you installed Python 3.8 32-bit, you should download ``numpy‑1.17.3+mkl‑cp38‑cp38‑win32.whl``.

For each downloaded file, run ``pip install <wheel file>``, e.g. ``pip install numpy‑1.17.3+mkl‑cp38‑cp38‑win32.whl``.

Be sure to install [numpy](https://numpy.org/) before [pandas](https://pandas.pydata.org/), as pandas depends on numpy.

## Configuring your firewall

Be sure to allow ``python`` to perform outbound connections

# Running

Run ``python _3GPP_Meeting_Helper.py`` to start the GUI. It should start by asking you whether you want to use an HTTP proxy or not.

It has been seen that sometimes some firewall issues happen, in which case the application may not run properly. In order to check whether Python has Internet access (not a proxy check, just whether it can do HTTP requests), you can run ``python`` and once you are in the Python prompt, use the following code:

```
import requests
requests.get('http://www.3gpp.org/ftp/Meetings_3GPP_SYNC/SA2/TdocsByAgenda.htm')
```

You should get a ``"(Response [200])`` string (maybe another response code if the file does not exist, which it did at the time this was written) and not a long list of exceptions. If the latter is the case, this means that the HTTP request could not reach the server.

# Usage

# Opening a specific TDoc

Select the meeting from the drop-down list, input the TDoc number (e.g. S2-2004143) and press the button to download and open the TDoc.

The TDoc will be searched for in the ``Docs`` folder, as well as in the local inbox if you find yourself in the meeting premises (detected by the ``10.10.10.XX`` IP address range allocated during the meeting).

Note: for eMeetings, you can also search for specific revisions, e.g. S2-2004143r01. In this case, the ``Docs/Revisions`` folder is used.
