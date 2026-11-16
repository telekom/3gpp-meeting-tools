# --- File: src/modules/meetings/core/session_plan_timezone.py ---
"""
Timezone engine for 3GPP Meeting Session Plans and World Clock.
Solves the corporate Austrian laptop clock problem (Europe/Vienna CET/CEST)
by strictly anchoring calculations in true UTC and translating between
Venue Local Time and Laptop Local Time.
"""
import json
import logging
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Optional, Tuple

try:
    from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
except ImportError:  # pragma: no cover - Python < 3.9 fallback
    ZoneInfo = None  # type: ignore[assignment]
    ZoneInfoNotFoundError = Exception  # type: ignore[assignment]

logger = logging.getLogger(__name__)

# Known 3GPP meeting host cities & countries with fallback standard/DST offsets (in minutes)
# Format: {city_key: (std_offset_min, dst_offset_min, dst_months_start, dst_months_end, tz_name)}
KNOWN_3GPP_CITIES: Dict[str, Tuple[int, int, int, int, str]] = {
    # --- Europe (CET/CEST) ---
    "prague": (60, 120, 3, 10, "Europe/Prague (CET/CEST)"),
    "vienna": (60, 120, 3, 10, "Europe/Vienna (CET/CEST)"),
    "innsbruck": (60, 120, 3, 10, "Europe/Vienna (CET/CEST)"),
    "saalfelden": (60, 120, 3, 10, "Europe/Vienna (CET/CEST)"),
    "voesendorf": (60, 120, 3, 10, "Europe/Vienna (CET/CEST)"),
    "gothenburg metropolitan area": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "gothenburg": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "stor-göteborg": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "goteborg": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "goterborg": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "gothenberg": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "stockholm": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "kista": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "malmo": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "malmö": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "lulea": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "sundsvall": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "haninge": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "lidingo": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "solna": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "ystad": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "bratislava": (60, 120, 3, 10, "Europe/Bratislava (CET/CEST)"),
    "maastricht": (60, 120, 3, 10, "Europe/Amsterdam (CET/CEST)"),
    "rotterdam": (60, 120, 3, 10, "Europe/Amsterdam (CET/CEST)"),
    "the hague": (60, 120, 3, 10, "Europe/Amsterdam (CET/CEST)"),
    "dubrovnik": (60, 120, 3, 10, "Europe/Zagreb (CET/CEST)"),
    "split": (60, 120, 3, 10, "Europe/Zagreb (CET/CEST)"),
    "zagreb": (60, 120, 3, 10, "Europe/Zagreb (CET/CEST)"),
    "valencia": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "barcelona": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "madrid": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "sevilla": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "seville": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "malaga": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "sitges": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "geneva": (60, 120, 3, 10, "Europe/Zurich (CET/CEST)"),
    "sophia-antipolis": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "sophia antipolis": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "st laurent du var": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "issy-les-moulineaux": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "nice": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "cannes": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "paris": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "biarritz": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "bordeaux": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "rennes": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "toulouse": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "montrouge": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "suresnes": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "versailles": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "valletta": (60, 120, 3, 10, "Europe/Malta (CET/CEST)"),
    "saint julian's": (60, 120, 3, 10, "Europe/Malta (CET/CEST)"),
    "st julian": (60, 120, 3, 10, "Europe/Malta (CET/CEST)"),
    "saint paul's bay": (60, 120, 3, 10, "Europe/Malta (CET/CEST)"),
    "malta": (60, 120, 3, 10, "Europe/Malta (CET/CEST)"),
    "monte carlo": (60, 120, 3, 10, "Europe/Monaco (CET/CEST)"),
    "rome": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "roma": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "milan": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "milano": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "turin": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "torino": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "florence": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "venice": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "sorrento": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "lecce": (60, 120, 3, 10, "Europe/Rome (CET/CEST)"),
    "berlin": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "munich": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "frankfurt": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "dusseldorf": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "bonn": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "mainz": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "dresden": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "erlangen": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "hannover": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "heidelberg": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "muenster": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "radebeul": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "bernried": (60, 120, 3, 10, "Europe/Berlin (CET/CEST)"),
    "warsaw": (60, 120, 3, 10, "Europe/Warsaw (CET/CEST)"),
    "wroclaw": (60, 120, 3, 10, "Europe/Warsaw (CET/CEST)"),
    "krakow": (60, 120, 3, 10, "Europe/Warsaw (CET/CEST)"),
    "budapest": (60, 120, 3, 10, "Europe/Budapest (CET/CEST)"),
    "belgrade": (60, 120, 3, 10, "Europe/Belgrade (CET/CEST)"),
    "ljubljana": (60, 120, 3, 10, "Europe/Ljubljana (CET/CEST)"),
    "portoroz": (60, 120, 3, 10, "Europe/Ljubljana (CET/CEST)"),
    "sarajevo": (60, 120, 3, 10, "Europe/Sarajevo (CET/CEST)"),
    "brussels": (60, 120, 3, 10, "Europe/Brussels (CET/CEST)"),
    "antwerp": (60, 120, 3, 10, "Europe/Brussels (CET/CEST)"),
    "bruges": (60, 120, 3, 10, "Europe/Brussels (CET/CEST)"),
    "copenhagen": (60, 120, 3, 10, "Europe/Copenhagen (CET/CEST)"),
    "aalborg": (60, 120, 3, 10, "Europe/Copenhagen (CET/CEST)"),
    "taastrup": (60, 120, 3, 10, "Europe/Copenhagen (CET/CEST)"),
    "oslo": (60, 120, 3, 10, "Europe/Oslo (CET/CEST)"),
    "trondheim": (60, 120, 3, 10, "Europe/Oslo (CET/CEST)"),
    "brunstad": (60, 120, 3, 10, "Europe/Oslo (CET/CEST)"),
    # --- Europe (EET/EEST & WET/WEST & TRT) ---
    "helsinki": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "espoo": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "oulu": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "tampere": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "turku": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "naantali": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "hämeenlinna": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "finland": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "athens": (120, 180, 3, 10, "Europe/Athens (EET/EEST)"),
    "bucharest": (120, 180, 3, 10, "Europe/Bucharest (EET/EEST)"),
    "sofia": (120, 180, 3, 10, "Europe/Sofia (EET/EEST)"),
    "tallinn": (120, 180, 3, 10, "Europe/Tallinn (EET/EEST)"),
    "riga": (120, 180, 3, 10, "Europe/Riga (EET/EEST)"),
    "vilnius": (120, 180, 3, 10, "Europe/Vilnius (EET/EEST)"),
    "limassol": (120, 180, 3, 10, "Asia/Nicosia (EET/EEST)"),
    "london": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "edinburgh": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "birmingham": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "brighton": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "bristol": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "bath": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "cambridge": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "cardiff": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "newbury": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "oxford": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "queensferry": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "swindon": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "warwick": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "east molesley": (0, 60, 3, 10, "Europe/London (GMT/BST)"),
    "dublin": (0, 60, 3, 10, "Europe/Dublin (GMT/IST)"),
    "cork": (0, 60, 3, 10, "Europe/Dublin (GMT/IST)"),
    "lisbon": (0, 60, 3, 10, "Europe/Lisbon (WET/WEST)"),
    "lisboa": (0, 60, 3, 10, "Europe/Lisbon (WET/WEST)"),
    "oporto": (0, 60, 3, 10, "Europe/Lisbon (WET/WEST)"),
    "povoa de varzim": (0, 60, 3, 10, "Europe/Lisbon (WET/WEST)"),
    "tenerife": (0, 60, 3, 10, "Atlantic/Canary (WET/WEST)"),
    "santa cruz": (0, 60, 3, 10, "Atlantic/Canary (WET/WEST)"),
    "istanbul": (180, 180, 0, 0, "Europe/Istanbul (TRT, UTC+3)"),
    # --- Africa ---
    "cape town": (120, 120, 0, 0, "Africa/Johannesburg (SAST, UTC+2)"),
    "capetown": (120, 120, 0, 0, "Africa/Johannesburg (SAST, UTC+2)"),
    # --- Asia ---
    "fukuoka": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "osaka": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "tokyo": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "kyoto": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "yokohama": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "kobe": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "nagoya": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "sapporo": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "okinawa": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "chiba": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "makuhari": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "kanazawa": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "miyazaki": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "miyasaki": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "abiko": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "hayama": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "yokosuka": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "japan": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "incheon": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "jeju": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "cheju": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "seoul": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "busan": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "pusan": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "gyeongju": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "kyongju": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "south korea": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "korea": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "wuhan": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "shanghai": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "beijing": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "shenzhen": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "chengdu": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "chongqing": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "guangzhou": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "hangzhou": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "nanjing": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "qingdao": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "sanya": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "suzhou": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "xi'an": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "xian": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "xiamen": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "zhuhai": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "dalian": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "da lian": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "dali": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "dongguan": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "guilin": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "harbin": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "hefei": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "kunming": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "changsha": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "zhangjiajie": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "china": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "hong kong": (480, 480, 0, 0, "Asia/Hong_Kong (HKT, UTC+8)"),
    "taipei": (480, 480, 0, 0, "Asia/Taipei (CST, UTC+8)"),
    "kaohsiung": (480, 480, 0, 0, "Asia/Taipei (CST, UTC+8)"),
    "singapore": (480, 480, 0, 0, "Asia/Singapore (SGT, UTC+8)"),
    "bengaluru": (330, 330, 0, 0, "Asia/Kolkata (IST, UTC+5:30)"),
    "bangalore": (330, 330, 0, 0, "Asia/Kolkata (IST, UTC+5:30)"),
    "new delhi": (330, 330, 0, 0, "Asia/Kolkata (IST, UTC+5:30)"),
    "hyderabad": (330, 330, 0, 0, "Asia/Kolkata (IST, UTC+5:30)"),
    "chennai": (330, 330, 0, 0, "Asia/Kolkata (IST, UTC+5:30)"),
    "kochi": (330, 330, 0, 0, "Asia/Kolkata (IST, UTC+5:30)"),
    "india": (330, 330, 0, 0, "Asia/Kolkata (IST, UTC+5:30)"),
    "bangkok": (420, 420, 0, 0, "Asia/Bangkok (ICT, UTC+7)"),
    # --- North & South America ---
    "calgary": (-420, -360, 3, 11, "America/Edmonton (MST/MDT)"),
    "vancouver": (-480, -420, 3, 11, "America/Vancouver (PST/PDT)"),
    "victoria": (-480, -420, 3, 11, "America/Vancouver (PST/PDT)"),
    "montreal": (-300, -240, 3, 11, "America/Toronto (EST/EDT)"),
    "quebec": (-300, -240, 3, 11, "America/Toronto (EST/EDT)"),
    "toronto": (-300, -240, 3, 11, "America/Toronto (EST/EDT)"),
    "reno": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "san diego": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "sparks": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "los angeles": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "san francisco": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "san franciso": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "san jose": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "seattle": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "redmond": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "las vegas": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "palm springs": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "palmsprings": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "newport beach": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "la jolla": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "anaheim": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "santa rosa": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "marina del rey": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "lake tahoe": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "spokane": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "denver": (-420, -360, 3, 11, "America/Denver (MST/MDT)"),
    "albuquerque": (-420, -360, 3, 11, "America/Denver (MST/MDT)"),
    "salt lake city": (-420, -360, 3, 11, "America/Denver (MST/MDT)"),
    "phoenix": (-420, -420, 0, 0, "America/Phoenix (MST, UTC-7)"),
    "scottsdale": (-420, -420, 0, 0, "America/Phoenix (MST, UTC-7)"),
    "chicago": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "austin": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "dallas": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "new orleans": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "nashville": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "san antonio": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "st louis": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "memphis": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "naperville": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "kansas city": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "texas": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "boston": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "baltimore": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "orlando": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "miami": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "west palm beach": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "marco island": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "fort lauderdale": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "jacksonville": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "atlanta": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "washington": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "new york": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "new jersey": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "jersey city": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "philadelphia": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "dulles": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "fairfax": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "virginia": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "morristown": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "east brunswick": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "alaska": (-540, -480, 3, 11, "America/Anchorage (AKST/AKDT)"),
    "hawaii": (-600, -600, 0, 0, "Pacific/Honolulu (HST, UTC-10)"),
    "maui": (-600, -600, 0, 0, "Pacific/Honolulu (HST, UTC-10)"),
    "kapolei": (-600, -600, 0, 0, "Pacific/Honolulu (HST, UTC-10)"),
    "puerto rico": (-240, -240, 0, 0, "America/Puerto_Rico (AST, UTC-4)"),
    "saint kitts": (-240, -240, 0, 0, "America/St_Kitts (AST, UTC-4)"),
    "oranjestad": (-240, -240, 0, 0, "America/Aruba (AST, UTC-4)"),
    "cancun": (-300, -300, 0, 0, "America/Cancun (EST, UTC-5)"),
    "puerto vallarta": (-360, -360, 0, 0, "America/Mexico_City (CST, UTC-6)"),
    "acapulco": (-360, -360, 0, 0, "America/Mexico_City (CST, UTC-6)"),
    "san jose del cabo": (-420, -420, 0, 0, "America/Mazatlan (MST, UTC-7)"),
    "rio de janeiro": (-180, -180, 0, 0, "America/Sao_Paulo (BRT, UTC-3)"),
    # --- Oceania ---
    "gold coast": (600, 600, 0, 0, "Australia/Brisbane (AEST, UTC+10)"),
    "brisbane": (600, 600, 0, 0, "Australia/Brisbane (AEST, UTC+10)"),
    "melbourne": (600, 660, 10, 4, "Australia/Melbourne (AEST/AEDT)"),
    "sydney": (600, 660, 10, 4, "Australia/Sydney (AEST/AEDT)"),
    "christchurch": (720, 780, 9, 4, "Pacific/Auckland (NZST/NZDT)"),
}

# Canonical IANA timezone ID extracted from KNOWN_3GPP_CITIES
CITY_TO_IANA_TZ: Dict[str, str] = {
    k: v[4].split(" ", 1)[0].strip() for k, v in KNOWN_3GPP_CITIES.items()
}

# Longest-first ordering prevents short keys (e.g. "san jose") from shadowing
# longer keys (e.g. "san jose del cabo").
_SORTED_CITY_KEYS = sorted(KNOWN_3GPP_CITIES.keys(), key=len, reverse=True)

MEETING_TZ_FILENAME = "meeting_timezone.json"


def format_utc_offset(offset_minutes: int) -> str:
    """Formats a UTC offset in minutes into 'UTC+HH:MM' or 'UTC-HH:MM' safely."""
    sign = "+" if offset_minutes >= 0 else "-"
    abs_min = abs(int(offset_minutes))
    hours, mins = divmod(abs_min, 60)
    return f"UTC{sign}{hours:02d}:{mins:02d}"


class SessionPlanTimezone:
    """Manages timezone calculations, dual clocks, per-meeting persistence, and UTC conversions."""

    MEETING_TZ_FILENAME = MEETING_TZ_FILENAME

    @classmethod
    def get_laptop_tz_offset_minutes(cls) -> int:
        """
        Calculates the user's laptop local UTC offset in minutes right now.
        For an Austrian corporate laptop, this returns +60 (CET in winter) or +120 (CEST in summer).
        """
        now = datetime.now()
        utc_now = datetime.now(timezone.utc).replace(tzinfo=None)
        diff = now - utc_now
        return int(round(diff.total_seconds() / 60))

    @classmethod
    def get_laptop_tz_label(cls) -> str:
        """User-friendly label for laptop local time."""
        mins = cls.get_laptop_tz_offset_minutes()
        return f"Austrian Laptop Time ({format_utc_offset(mins)})"

    @classmethod
    def extract_tz_from_agenda_text(cls, text: str) -> Optional[Tuple[int, str]]:
        """
        Tier 1: Parses Agenda document deadlines table text for Local time vs UTC.
        Example: '| Start of meeting | Monday 12 October 2026 | 0900 Local time | 0700 UTC |'
        Computes offset = 09:00 - 07:00 = +120 minutes (UTC+02:00).
        """
        if not text:
            return None

        patterns = [
            r"(\d{3,4})\s*(?:hours|hrs)?\s*Local\s*time.*?(\d{3,4})\s*(?:hours|hrs)?\s*UTC",
            r"(\d{1,2}:\d{2})\s*Local\s*time.*?(\d{1,2}:\d{2})\s*UTC",
        ]

        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE | re.DOTALL)
            if m:
                local_raw, utc_raw = m.group(1), m.group(2)
                local_min = cls._time_str_to_minutes(local_raw)
                utc_min = cls._time_str_to_minutes(utc_raw)
                if local_min is not None and utc_min is not None:
                    diff_min = local_min - utc_min
                    # Handle day wrapping (e.g. 01:00 Local vs 23:00 UTC)
                    if diff_min > 720:
                        diff_min -= 1440
                    elif diff_min < -720:
                        diff_min += 1440

                    label = f"{format_utc_offset(diff_min)} (from Agenda Deadlines)"
                    logger.info(f"[SessionPlanTZ] Auto-detected meeting timezone from Agenda: {diff_min}m ({label})")
                    return (diff_min, label)

        return None

    @classmethod
    def _match_city_key(cls, location_str: str) -> Optional[str]:
        if not location_str:
            return None
        loc_lower = location_str.strip().lower()
        if not loc_lower:
            return None

        # 1. Word-boundary match (longest key first)
        for city_key in _SORTED_CITY_KEYS:
            pattern = r"(?<!\w)" + re.escape(city_key) + r"(?!\w)"
            if re.search(pattern, loc_lower):
                return city_key

        # 2. Substring fallback (longest key first)
        for city_key in _SORTED_CITY_KEYS:
            if city_key in loc_lower:
                return city_key

        return None

    @classmethod
    def _parse_reference_utc(cls, meeting_date_str: str = "") -> datetime:
        """Parses a meeting date string into a reference UTC datetime, or returns current UTC."""
        now_utc = datetime.now(timezone.utc)
        if not meeting_date_str:
            return now_utc

        s = meeting_date_str.strip()
        m_iso = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
        if m_iso:
            try:
                y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
                return datetime(y, m, d, 12, 0, tzinfo=timezone.utc)
            except ValueError:
                pass

        month_names = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
        m_mon = re.search(
            r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b",
            s,
            re.I,
        )
        if m_mon:
            month = month_names.index(m_mon.group(1)[:3].lower()) + 1
            m_year = re.search(r"\b(19\d{2}|20\d{2})\b", s)
            year = int(m_year.group(1)) if m_year else now_utc.year
            return datetime(year, month, 15, 12, 0, tzinfo=timezone.utc)

        m_num = re.search(r"[-/](\d{1,2})[-/]", s)
        if m_num:
            try:
                month = max(1, min(12, int(m_num.group(1))))
                return datetime(now_utc.year, month, 15, 12, 0, tzinfo=timezone.utc)
            except ValueError:
                pass

        return now_utc

    @classmethod
    def resolve_iana_offset_minutes(
        cls, iana_tz: str, ref_utc: Optional[datetime] = None
    ) -> Optional[Tuple[int, str]]:
        """Resolves an IANA timezone string to (offset_minutes, tz_abbr) using Python zoneinfo."""
        if not iana_tz or ZoneInfo is None:
            return None
        try:
            zi = ZoneInfo(iana_tz)
            utc_dt = ref_utc or datetime.now(timezone.utc)
            aware_dt = utc_dt.astimezone(zi)
            utc_offset = aware_dt.utcoffset()
            if utc_offset is None:
                return None
            offset_min = int(round(utc_offset.total_seconds() / 60))
            abbr = aware_dt.tzname() or ""
            return offset_min, abbr
        except ZoneInfoNotFoundError:
            return None
        except Exception as e:
            logger.debug(f"[SessionPlanTZ] ZoneInfo resolution failed for '{iana_tz}': {e}")
            return None

    @classmethod
    def lookup_city_timezone_details(
        cls, location_str: str, meeting_date_str: str = ""
    ) -> Optional[Tuple[int, str, str]]:
        """
        Matches meeting location string against known 3GPP venue cities/countries.
        Uses Python zoneinfo (exact IANA DST rules) when available, falling back to
        static month-based DST rules if tzdata is missing.
        Returns (offset_minutes, label, iana_tz) or None.
        """
        matched_city = cls._match_city_key(location_str)
        if not matched_city:
            return None

        std_min, dst_min, dst_start, dst_end, tz_name = KNOWN_3GPP_CITIES[matched_city]
        iana_tz = CITY_TO_IANA_TZ.get(matched_city, "")
        ref_utc = cls._parse_reference_utc(meeting_date_str)

        # 1. Preferred: exact IANA zoneinfo evaluation (handles regional DST law changes)
        if iana_tz:
            zi_res = cls.resolve_iana_offset_minutes(iana_tz, ref_utc=ref_utc)
            if zi_res is not None:
                chosen_offset, abbr = zi_res
                suffix = f"{iana_tz} - {abbr}" if abbr and not abbr.startswith(("+", "-")) else iana_tz
                label = f"{format_utc_offset(chosen_offset)} ({suffix})"
                return (chosen_offset, label, iana_tz)

        # 2. Fallback: static month-based DST heuristic
        month = ref_utc.month
        is_dst = False
        if dst_start > 0 and dst_end > 0:
            if dst_start < dst_end:
                is_dst = dst_start <= month <= dst_end
            else:
                # Southern hemisphere DST (spans year end)
                is_dst = month >= dst_start or month <= dst_end

        chosen_offset = dst_min if is_dst else std_min
        label = f"{format_utc_offset(chosen_offset)} ({tz_name})"
        logger.debug(f"[SessionPlanTZ] Fallback resolved '{matched_city}' to: {chosen_offset}m ({label})")
        return (chosen_offset, label, iana_tz)

    @classmethod
    def lookup_city_timezone(cls, location_str: str, meeting_date_str: str = "") -> Optional[Tuple[int, str]]:
        """
        Tier 2: Matches meeting location string against known 3GPP venue cities.
        Accounts for Daylight Saving Time based on meeting date (or current date).
        """
        details = cls.lookup_city_timezone_details(location_str, meeting_date_str=meeting_date_str)
        if details is None:
            return None
        offset_min, label, _iana = details
        return (offset_min, label)

    @classmethod
    def load_meeting_folder_timezone(cls, meeting_dir: Path) -> Optional[dict]:
        """Reads `<meeting_dir>/meeting_timezone.json` if it exists and is valid."""
        if not meeting_dir:
            return None
        tz_file = Path(meeting_dir) / MEETING_TZ_FILENAME
        if not tz_file.is_file():
            return None
        try:
            data = json.loads(tz_file.read_text(encoding="utf-8"))
            if not isinstance(data, dict) or data.get("tz_offset_minutes") is None:
                return None
            return {
                "wg": str(data.get("wg", "")),
                "meeting_number": str(data.get("meeting_number", "")),
                "location": str(data.get("location", "")),
                "tz_offset_minutes": int(data["tz_offset_minutes"]),
                "tz_label": str(data.get("tz_label", "")),
                "iana_tz": str(data.get("iana_tz", "")),
                "source": str(data.get("source", "meeting_folder")),
                "updated_at": str(data.get("updated_at", "")),
            }
        except Exception as e:
            logger.debug(f"[SessionPlanTZ] Failed to load {tz_file}: {e}")
            return None

    @classmethod
    def save_meeting_folder_timezone(
        cls,
        meeting_dir: Path,
        tz_offset_minutes: int,
        tz_label: str,
        source: str = "manual_override",
        location: str = "",
        wg: str = "",
        meeting_number: str = "",
        iana_tz: str = "",
    ) -> bool:
        """
        Persists confirmed/overridden timezone configuration inside `<meeting_dir>/meeting_timezone.json`
        and synchronizes `Agenda/session_plans.db` (`meeting_meta`) if that database exists.
        """
        if not meeting_dir:
            return False
        mtg_path = Path(meeting_dir)
        tz_file = mtg_path / MEETING_TZ_FILENAME
        payload = {
            "wg": str(wg or ""),
            "meeting_number": str(meeting_number or ""),
            "location": str(location or ""),
            "tz_offset_minutes": int(tz_offset_minutes),
            "tz_label": str(tz_label or format_utc_offset(int(tz_offset_minutes))),
            "iana_tz": str(iana_tz or ""),
            "source": str(source or "manual_override"),
            "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        try:
            mtg_path.mkdir(parents=True, exist_ok=True)
            tmp_file = tz_file.with_suffix(".json.tmp")
            tmp_file.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
            os.replace(tmp_file, tz_file)
        except Exception as e:
            logger.warning(f"[SessionPlanTZ] Could not save {tz_file}: {e}")
            return False

        # Keep existing Agenda/session_plans.db synchronized when present
        sp_db_path = mtg_path / "Agenda" / "session_plans.db"
        if sp_db_path.is_file() and meeting_number:
            try:
                from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase

                sp_db = SessionPlanDatabase(sp_db_path)
                sp_db.set_meeting_meta(
                    meeting_number=str(meeting_number),
                    location=str(location or ""),
                    tz_offset_minutes=int(tz_offset_minutes),
                    tz_name=payload["tz_label"],
                    tz_confirmed=1,
                )
            except Exception as e:
                logger.debug(f"[SessionPlanTZ] Could not sync session_plans.db: {e}")

        return True

    @classmethod
    def clear_meeting_folder_timezone(cls, meeting_dir: Path) -> bool:
        """Removes `<meeting_dir>/meeting_timezone.json` if it exists."""
        if not meeting_dir:
            return False
        tz_file = Path(meeting_dir) / MEETING_TZ_FILENAME
        if tz_file.is_file():
            try:
                tz_file.unlink()
                return True
            except Exception as e:
                logger.debug(f"[SessionPlanTZ] Could not delete {tz_file}: {e}")
                return False
        return False

    @staticmethod
    def _time_str_to_minutes(s: str) -> Optional[int]:
        """Converts '0900', '900', '09:00', '9:00' to minutes since midnight."""
        s = re.sub(r"[^\d:]", "", s.strip())
        if ":" in s:
            parts = s.split(":")
            return int(parts[0]) * 60 + int(parts[1])
        if len(s) in (3, 4):
            hh = int(s[:-2])
            mm = int(s[-2:])
            return hh * 60 + mm
        return None

    @classmethod
    def get_venue_current_datetime(cls, meeting_tz_offset_min: int) -> datetime:
        """
        Converts the current physical instant into Venue Local datetime,
        completely immune to the laptop clock's local timezone.
        """
        utc_now = datetime.now(timezone.utc)
        venue_tz = timezone(timedelta(minutes=meeting_tz_offset_min))
        return utc_now.astimezone(venue_tz)

    @classmethod
    def format_dual_time_range(
        cls, start_hhmm: str, end_hhmm: str, meeting_tz_offset_min: int
    ) -> str:
        """
        Formats a time range for display.
        If laptop timezone differs from meeting venue, appends the Austrian laptop time in parentheses.
        Example: '09:00 – 10:30 (02:00 – 03:30 Vienna)'
        """
        laptop_offset_min = cls.get_laptop_tz_offset_minutes()
        offset_diff = laptop_offset_min - meeting_tz_offset_min

        venue_str = f"{start_hhmm} – {end_hhmm}"
        if offset_diff == 0:
            return venue_str

        # Calculate corresponding Austrian laptop times
        s_min = cls._time_str_to_minutes(start_hhmm)
        e_min = cls._time_str_to_minutes(end_hhmm)
        if s_min is None or e_min is None:
            return venue_str

        lap_s = (s_min + offset_diff) % 1440
        lap_e = (e_min + offset_diff) % 1440

        lap_str = f"{lap_s // 60:02d}:{lap_s % 60:02d} – {lap_e // 60:02d}:{lap_e % 60:02d} Vienna"
        return f"{venue_str} ({lap_str})"

    @classmethod
    def venue_local_to_utc_dt(
        cls, date_str: str, time_hhmm: str, meeting_tz_offset_min: int
    ) -> Optional[datetime]:
        """
        Converts a Venue Local date ('2026-10-12') and time ('09:00') into an aware UTC datetime object.
        Essential for generating accurate RFC 5545 .ics calendar entries.
        """
        try:
            m = cls._time_str_to_minutes(time_hhmm)
            if m is None:
                return None
            hh = m // 60
            mm = m % 60

            # Parse date
            d = datetime.strptime(date_str, "%Y-%m-%d").date()
            venue_dt = datetime(d.year, d.month, d.day, hh, mm)
            venue_tz = timezone(timedelta(minutes=meeting_tz_offset_min))
            venue_aware = venue_dt.replace(tzinfo=venue_tz)
            return venue_aware.astimezone(timezone.utc)
        except Exception as e:
            logger.debug(f"[SessionPlanTZ] Error converting {date_str} {time_hhmm} to UTC: {e}")
            return None
