import os
import re
import win32ui
import dde
import sys
import time
import datetime
import win32com.client
from termcolor import colored

DATA_MAP = [
    ("sat_name", '^SN"(.*)"[ |\r\n]'),
    ("sat_azimuth", ' AZ(\\S*)[ |\r\n]'),     # range form 0.0 to 360.0
    ("sat_elevation", ' EL(\\S*)[ |\r\n]'),   # range from -90.0 to 90.0
    ("sat_downlink", ' DN(\\S*)[ |\r\n]'),    # with doppler, [Hz] in range from 0 to N
    ("sat_uplink", ' UP(\\S*)[ |\r\n]'),      # with doppler, [Hz] in range from 0 to N
    ("sat_distance", ' RA(\\S*)[ |\r\n]'),    # km
    ("sat_velocity", ' RR(\\S*)[ |\r\n]'),    # km/s, negative=approaching positive=going-away
    ("sat_longitude", ' LO(\\S*)[ |\r\n]'),   # range from -180.0000 to 180.000 (W to E)
    ("sat_latitude", ' LA(\\S*)[ |\r\n]'),    # range from -90.0000 to 90.000 (S to N)
    ("sat_altitude", ' AL(\\S*)[ |\r\n]'),    # km
    ("timestamp_utc", ' TU(\\S*)[ |\r\n]'),   # YYYYMMDDhhmmss
    ("timestamp_local", ' TL(\\S*)[ |\r\n]'), # YYYYMMDDhhmmss
]
COMPASS = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE",
           "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
OMNIRIG_VFOA = 2048
OMNIRIG_VFOB = 4096
OMNIRIG_SPLITON = 32768
OMNIRIG_SPLITOFF = 65536
OMNIRIG_MODEFM = 1073741824
FREQ_MAX_DELTA = 2000

# Global variables :)
FREQ_CURR_DL = 0  # Last set DL freq
FREQ_CURR_UL = 0  # Last set UL freq
RIG = False


def rig_init():
    global RIG
    RIG = win32com.client.Dispatch("Omnirig.OmnirigX")
    RIG.Rig1.Vfo = OMNIRIG_VFOA
    RIG.Rig1.Mode = OMNIRIG_MODEFM
    RIG.Rig1.Vfo = OMNIRIG_VFOB
    RIG.Rig1.Mode = OMNIRIG_MODEFM
    RIG.Rig1.Vfo = OMNIRIG_VFOA
    RIG.Rig1.Split = OMNIRIG_SPLITON


def rig_get_data():
    global RIG
    rig_type = RIG.Rig1.RigType
    rig_status = RIG.Rig1.StatusStr
    rig_data = {"rig_type": rig_type, "rig_status": rig_status}
    return rig_data


def rig_get_freq():
    global RIG
    return RIG.Rig1.Freq


def rig_set_freq(vfo, freq):
    global RIG
    RIG.Rig1.Vfo = vfo
    RIG.Rig1.Freq = freq
    RIG.Rig1.Vfo = OMNIRIG_VFOA


def deg_to_compass(angle):
    val = int((angle/22.5)+.5)
    return COMPASS[(val % 16)]


def adjust_freq(sat_data):
    global FREQ_CURR_DL
    global FREQ_CURR_UL
    exact_dl = float(sat_data["sat_downlink"])
    ecact_ul = float(sat_data["sat_uplink"])

    display_freq = int(rig_get_freq())
    if abs(display_freq-int(FREQ_CURR_UL)) <= FREQ_MAX_DELTA:
        # We're on VfoB
        # We're transmitting
        rig_set_freq(vfo=OMNIRIG_VFOB, freq=ecact_ul)
        FREQ_CURR_UL = ecact_ul
    else:
        # We assume we're on VfoA
        # We're receiving or not yet initialized
        rig_set_freq(vfo=OMNIRIG_VFOA, freq=exact_dl)
        FREQ_CURR_DL = exact_dl
    return


def handle_data(sat_data_raw, rig_data):
    sat_data = parse_sat_data(sat_data_raw)
    display_data(sat_data, rig_data)
    adjust_freq(sat_data)


def parse_sat_data(data):
    parsed_dict = {}
    for value in DATA_MAP:
        try:
            parsed_dict[value[0]] = re.findall(value[1], data)[0]
        except Exception:
            parsed_dict[value[0]] = False
    return (parsed_dict)


def display_data(sat_data, rig_data):
    os.system('cls')
    try:
        timestamp_local = datetime.datetime.strptime(
            sat_data['timestamp_local'], '%Y%m%d%H%M%S')
        timestamp_utc = datetime.datetime.strptime(
            sat_data['timestamp_utc'], '%Y%m%d%H%M%S')
        print("Timestamp :", timestamp_local, "local time")
        print("Timestamp :", timestamp_utc, "UTC")
        print("Radio     :", rig_data['rig_type'], end=' ')
        if rig_data['rig_status'] != "On-line":
            print(colored("off-line", "red"))
        else:
            print("on-line")
        print("Satellite :", sat_data['sat_name'])
        print("Obs. PoV sat. Elevation  :", int(
            float(sat_data['sat_elevation'])), "deg", end='')
        if float(sat_data['sat_elevation']) > 0:
            print(" (above the horizon, visible)")
        else:
            print(colored(" (below the horizon, not visible)", "red"))
        print("Obs. PoV sat. Azimuth    :", int(
            float(sat_data['sat_azimuth'])), "deg", end=' ')
        print("("+deg_to_compass(float(sat_data['sat_azimuth']))+")")
        print("Obs. PoV sat. Distance   :", int(
            float(sat_data['sat_distance'])), "km")
        print("Obs. PoV sat. Velocity   :", int(
            float(sat_data['sat_velocity'])*60*60), "km/h", end='')
        if float(sat_data['sat_velocity']) < 0:
            print(" (approaching)")
        else:
            print(" (going away)")
        print("Obs. PoV sat. Freq. up   :", sat_data['sat_uplink'], "Hz")
        print("Obs. PoV sat. Freq. down :", sat_data['sat_downlink'], "Hz")
    except Exception:
        print("Broken data!")


if __name__ == "__main__":
    os.system('mode con lines=13 cols=70')
    rig = False
    server = dde.CreateServer()
    server.Create("Orbitron2Omnirig")
    conversation = dde.CreateConversation(server)
    while True:
        try:
            # Initial try-to-connect loop
            while True:
                try:
                    conversation.ConnectTo("Orbitron", "Tracking")
                    rig_init()
                    break
                except Exception as e:
                    print("DDE conversation error: ",
                          e, " - Retrying in 1 sec...")
                    time.sleep(1)
            # Handle data loop
            while True:
                sat_data_raw = conversation.Request("TrackingDataEx")
                rig_data = rig_get_data()
                handle_data(sat_data_raw, rig_data)
                time.sleep(1)

        except KeyboardInterrupt:
            sys.exit(0)
        except Exception as e:
            print(e)
            input()
            sys.exit(1)
