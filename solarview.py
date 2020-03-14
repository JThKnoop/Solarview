"""
---------------------------

File:    solarview.py
         Produces a per year heatmap of solar PV energy production
         using the Growatt server data

         Reads the production data from the Growatt server (growatt.server.com)
         using the GrowattShinePhone api
         stores them locally.

         Needs a solarview.ini file with the following contents:
         
         [ini]
         username=YourUserName
         password=YourPassword
         pickle_dir="./"
         pickle_template="solarviewdata_????.pkl"

         Dependencies:
         requests
         
         Uses:
         GrowattApi: https://github.com/Sjord/growatt_api_client
         included in this file
         

Date:    2018-06-09
Author:  Jan Knoop, Ruud van der Ham

Version    Date        Change
   0.27    2020-03-13  Introduces pathlib, object g to avoid globals
                       Imports changed
                       Removed os, glob
   0.26    2020-02-16  Handles no connection
   0.25    2020-02-15  Corrects error in yearproduction
                       Remove ShinePhoneSample
   0.24    2020-02-14  Loads yeartotal from server, introduces ini-file
                       Format pickle-file changed
   0.23    2020-01-24  Saves and restores cache in local pickle-file
   0.22    2020-01-17  Future compatability
                       Default .png-filename
   0.21    2020-01-10  Working on Ubuntu: font + wait_visibility
   0.20    2020-01-06  Formatted and several minor changes
   0.19    2020-01-02  Select previous year from menu
   0.18    2020-01-01  Show previous years
   0.17    2019-06-02  Include ShinephoneApi from Sjord
                       Remove local cached history file
   0.16    2019-04-20  Use month-data
   0.15    2019-04-16  Read Growatt ShineAPI (Sjord)
   0.14    2019-03-29  Orientation grid switched
   0.13    2019-01-28  New parse algorithm
   0.12    2019-01-22  Only parse 4 columns
   0.11    2019-01-19  Aggregate to one log-file
   0.10    2019-01-19  Drawing in PIL, use general exception
    0.9    2019-01-05  Monsters --> Samples
    0.8    2018-12-30  Draw via pil
    0.7    2018-11-17  History-file format changed
    0.6    2018-07-03  Create legend
    0.5    2018-07-01  Create PIL image
    0.4    2018-06-19  Print labels
    0.3    2018-06-13  Even cleaner version
    0.2    2018-06-11  Clean version
    0.1    2018-06-09  Initial version
---------------------------

"""
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog

import datetime as dt
from PIL import Image, ImageDraw, ImageFont, ImageTk
import calendar

import pickle
import bz2
from pathlib import Path

from enum import IntEnum
import hashlib
import requests

import configparser

debug = False


class g:
    inifilename = "solarview.ini"


"""
Growatt Shinephone login data
"""


def hash_password(password):
    """
    Normal MD5, except add c if a byte of the digest is less than 10.
    """
    password_md5 = hashlib.md5(password.encode("utf-8")).hexdigest()
    for i in range(0, len(password_md5), 2):
        if password_md5[i] == "0":
            password_md5 = password_md5[0:i] + "c" + password_md5[i + 1 :]
    return password_md5


class Timespan(IntEnum):
    day = 1
    month = 2
    year = 3
    total = 4

    def format_date(self, date):
        if self == Timespan.day:
            return date.strftime("%Y-%m-%d")
        elif self == Timespan.month:
            return date.strftime("%Y-%m")
        elif self == Timespan.year:
            return date.strftime("%Y")
        elif self == Timespan.total:
            return ""
        else:
            raise ValueError(self)


class GrowattApiError(RuntimeError):
    pass


class LoginError(GrowattApiError):
    pass


class GrowattApi:
    server_url = "https://server.growatt.com/"

    def __init__(self):
        self.session = requests.Session()
        self.logged_in = False

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if self.logged_in:
            self.logout()

    def get_url(self, page):
        return self.server_url + page

    def login(self, username, password):
        """
        Log in to the Growatt server, or raise an exception if this fails.
        """
        password_md5 = hash_password(password)
        response = self.session.post(self.get_url("LoginAPI.do"), data={"userName": username, "password": password_md5})
        try:
            result = self._back_success_response(response)
            self.logged_in = True
            return result
        except GrowattApiError:
            raise LoginError

    def plant_list(self):
        """
        Retrieve all plants belonging to the current user.
        """
        response = self.session.get(self.get_url("PlantListAPI.do"), allow_redirects=False)
        return self._back_success_response(response)

    def plant_detail(self, plant_id, timespan, date):
        """
        Return amount of power generated for the given timespan.
        * Timespan.day : power on each half hour of the day.
        * Timespan.month : power on each day of the month.
        * Timespan.year: power on each month of the year.
        * Timespan.total: power on each year. `date` parameter is ignored.
        """
        assert timespan in Timespan
        date_str = timespan.format_date(date)

        response = self.session.get(
            self.get_url("PlantDetailAPI.do"), params={"plantId": plant_id, "type": timespan.value, "date": date_str}
        )
        return self._back_success_response(response)

    def new_plant_detail(self, plant_id, timespan, date):
        """
        Return amount of power generated for the given timespan.
        * Timespan.day : power on each five minutes of the day.
        * Timespan.month : power on each day of the month.
        * Timespan.year: power on each month of the year.
        * Timespan.total: power on each year. `date` parameter is ignored.
        """
        assert timespan in Timespan
        date_str = timespan.format_date(date)

        response = self.session.get(
            self.get_url("newPlantDetailAPI.do"), params={"plantId": plant_id, "type": timespan.value, "date": date_str}
        )
        return self._back_success_response(response)

    def get_user_center_energy_data(self):
        """
        Get overall data including:
        * powerValue - current power in Watt
        * todayValue - power generated today
        """
        response = self.session.post(
            self.get_url("newPlantAPI.do"), params={"action": "getUserCenterEnertyData"}, data={"language": 1}  # sic
        )
        return response.json()

    def logout(self):
        self.session.get(self.get_url("logout.do"))
        self.logged_in = False

    def _back_success_response(self, response):
        """
        Check and return the response, where we expect a "back" key with a
        "success" item.
        """
        if response.status_code != 200:
            raise GrowattApiError("Request failed: %s" % response)
        data = response.json()
        result = data["back"]
        if "success" in result and result["success"]:
            return result
        raise GrowattApiError()


"""
determine difference between two dates in iso8601 format
"""


def isodate_diff(d1_str, d2_str):
    f = "%Y-%m-%d"
    d1 = dt.datetime.strptime(d1_str, f)
    d2 = dt.datetime.strptime(d2_str, f)
    return (d1 - d2).days


"""
determine daytime in minutes
timestr="nn:nn:nn" or ""nn:nn"
"""


def isotime_to_m(timestr):
    m = [int(x) for x in timestr.split(":")]
    result = m[0] * 60 + m[1]
    if len(m) > 2:  # includes seconds
        result += m[2] / 60
    return result


"""
determine number of days in specified year
"""


def days_in_year(year):
    return (dt.datetime(year + 1, 1, 1) - dt.datetime(year, 1, 1)).days


"""
for debugging purposes
"""


def nowstr():
    return dt.datetime.now().strftime("%H:%M:%S")


"""
daterange start - end (inclusive)
"""


def daterange(start_date, end_date):
    for n in range(int((end_date - start_date).days) + 1):
        yield start_date + dt.timedelta(n)


class ShinePhoneDayData:
    def __init__(self, datestr, todayenergy):
        self.datestr = datestr
        self.todayenergy = todayenergy
        self.samples = {}  # timestampstr: actualpower


class GrowattServerData:
    """
    Read data of this year:
    - first check if local pickle file is present
    - if not or file not complete try to read from servers using ShinePhoneApi
      using Growatt api from Sjord
    """

    def __init__(self, year=None, setprogress=None):
        self.setprogress = setprogress

        self.yearsavailableonserver = {}

        now = dt.datetime.now()

        self.days = {}  # dict of ShinePhoneDayData()

        self.year_complete = False

        if year is None:
            self.year = now.year  #
            end_date = now
        else:
            self.year = year
            if year == now.year:
                end_date = now
            else:
                end_date = dt.datetime(self.year, 12, 31)

        if self.load_from_picklefile(self.year):
            self.plant_id = "plant_id"
            self.plant_name = "plant_name"

            """ determine which days to read """
            if len(self.days) > 0:
                start_date = dt.datetime.strptime(max(self.days.keys()), "%Y-%m-%d")  # string
            else:
                start_date = dt.datetime(self.year, 1, 1)

            if now.year == year:
                end_date = now
            else:
                end_date = dt.datetime(year, 12, 31)
        else:
            start_date = dt.datetime(year, 1, 1)

        if not self.year_complete:
            if self.downloadgrowattdata(start_date, end_date):
                if start_date.year < now.year:
                    self.year_complete = True
                self.dump_to_picklefile(self.year)

        self.yearsavailable = self.yearsavailablelocally()
        for year in self.yearsavailableonserver:
            if year not in self.yearsavailable:
                self.yearsavailable.append(year)
        self.yearsavailable.sort()

    def downloadgrowattdata(self, start_date, end_date):
        if debug:
            print("downloadgrowattdata: {} {}".format(start_date, end_date))

        try:
            with GrowattApi() as gwa:

                gwa.login(g.username, g.password)
                self.plant_info = gwa.plant_list()
                if debug:
                    print("**plant_info**", self.plant_info)
                self.plant_id = self.plant_info["data"][0]["plantId"]
                self.plant_name = ""

                """ determine for which years serverdata are available """
                plant_detail = gwa.new_plant_detail(self.plant_id, Timespan.total, None)
                if debug:
                    print("**plant_detail**", plant_detail)
                self.yearsavailableonserver = {
                    int(y): float(plant_detail["data"][y])
                    for y in plant_detail["data"]
                    if float(plant_detail["data"][y]) > 0
                }

                if debug:
                    print("yearsavailable on server", self.yearsavailableonserver)

                if start_date.year in self.yearsavailableonserver:
                    self.yearproduction = self.yearsavailableonserver[start_date.year]
                else:
                    self.yearproduction = 0.0

                """ prepare to show progress  """
                nr_of_days = (end_date - start_date).days + 1
                daycount = 0
                for d in daterange(start_date, end_date):
                    if self.setprogress is not None:
                        daycount += 1
                        self.setprogress(int(100 * daycount / nr_of_days))

                    plant_detail = gwa.new_plant_detail(self.plant_id, Timespan.day, d)
                    if debug:
                        print("**plant_detail**", plant_detail)
                    plantdata = plant_detail["plantData"]
                    if self.plant_name != plantdata["plantName"]:
                        self.plant_name = plantdata["plantName"]

                    datestr = d.strftime("%Y-%m-%d")
                    if datestr not in self.days:
                        self.days[datestr] = ShinePhoneDayData(datestr=datestr, todayenergy=0)

                    data = plant_detail["data"]
                    for datetimestampstr in data:
                        timestampstr = datetimestampstr.split(" ")[1]  # only time
                        actualpower = float(data[datetimestampstr])
                        self.days[datestr].samples[timestampstr] = actualpower

                    #                    self.days[datestr].samples.sort(key=lambda x: x.timestampstr)

                    if debug:
                        for ts in sorted(self.days[datestr].samples):
                            print(ts, self.days[datestr].samples[ts])

                """
                Read daily production (by reading monthly data)
                """
                months = []
                for d in daterange(start_date, end_date):
                    if not ((d.year, d.month) in months):
                        months.append((d.year, d.month))
                if debug:
                    print("months", months)

                for m in months:
                    monthstr = "{:4}-{:02}-".format(m[0], m[1])
                    firstday = dt.datetime(int(m[0]), int(m[1]), 1)
                    plantdetail_month = gwa.new_plant_detail(self.plant_id, Timespan.month, firstday)
                    if debug:
                        print(plantdetail_month)
                    monthdata = plantdetail_month["data"]
                    if debug:
                        print("monthdata", monthdata)

                    for day in monthdata.keys():
                        etoday = float(monthdata[day])
                        datestr = monthstr + day

                        if not (datestr in self.days):
                            self.days[datestr] = ShinePhoneDayData(datestr=datestr, todayenergy=0)
                        self.days[datestr].todayenergy = etoday

            result = True

        except GrowattApiError:
            tk.messagebox.showinfo("Connection Error", "No connection to Growatt servers")
            result = False

        except requests.exceptions.ConnectionError:
            tk.messagebox.showinfo("Connection Error", "No connection to Growatt servers")
            result = False

        except LoginError:
            tk.messagebox.showinfo("Login Error", "Username / password not correct")
            result = False

        return result  # True means data has been received from server

    """  Determine years available in local datafiles """

    def yearsavailablelocally(self):
        files = g.pickle_dir.glob(g.pickle_template)
        years = []
        for f in files:
            yearstr = "".join(cf for cf, cd in zip(f.name, g.pickle_template) if cd == "?")
            try:
                year = int(yearstr)
                years.append(year)
            except ValueError:
                pass
        return years

    def dump_to_picklefile(self, year):
        filename = Path(str(g.pickle_dir / g.pickle_template).replace("????", str(year)))
        with bz2.open(filename, "wb") as f:
            pickle.dump((self.year_complete, self.yearproduction, self.days), f)
        return True

    def load_from_picklefile(self, year):
        filename = Path(str(g.pickle_dir / g.pickle_template).replace("????", str(year)))
        if filename.exists():
            with bz2.open(filename, "rb") as f:
                self.year_complete, self.yearproduction, self.days = pickle.load(f)
            return True
        else:
            return False


class Projection:
    def __init__(self):
        self.height = 800
        self.topmargin = 50
        self.bottommargin = 50
        self.leftmargin = 50
        self.rightmargin = 50
        self.linewidth = 2
        self.pixels_per_day = self.linewidth + 1
        self.width = (self.linewidth + 1) * 365 + self.leftmargin + self.rightmargin

        self.pixels_per_kwh = 6
        self.pixels_per_hour = 30

        self.fiveoclockbase = self.height - self.bottommargin - 35 * self.pixels_per_kwh

        self.power_to_color_table = (
            (0, "#C0C0C0"),  # gray75
            (500, "#87CEFA"),  # light sky blue
            (1000, "#00FF00"),  # green
            (2000, "#FFFF00"),  # yellow
            (3000, "#FFA500"),  # orange
            (4000, "#FF0000"),  # red
            (float("inf"), "#B00000"),  # dark red
        )

    def power_to_color(self, power):
        for (upperbound, color) in self.power_to_color_table:
            if power <= upperbound:
                return color

    def power_legend(self):
        result = []
        lowerbound = None
        for (upperbound, color) in self.power_to_color_table:
            if lowerbound is None:
                result.append((color, "0 W"))
            elif upperbound == float("inf"):
                result.append((color, "> {} W".format(lowerbound)))
            else:
                result.append((color, "{} - {} W".format(lowerbound, upperbound)))
            lowerbound = upperbound
        return result


class YearSelector(simpledialog.Dialog):
    def __init__(self, parent, years):
        """ Init progress window """
        self.parent = parent
        tk.Toplevel.__init__(self, master=parent)
        self.year = 0

        """ Create progress window """
        self.focus_set()  # set focus on the ProgressWindow
        self.grab_set()  # make a modal window, so all events go to the ProgressWindow
        self.transient(self.master)  # show only one window in the task bar
        #
        self.title("Select year to plot")
        self.resizable(False, False)  # window is not resizable
        # self.close gets fired when the window is destroyed
        self.protocol("WM_DELETE_WINDOW", self.close)
        # Set proper position over the parent window
        self.geometry("200x200+100+100")
        self.bind("<Escape>", self.close)  # cancel progress when <Escape> key is pressed

        self.lbx = tk.Listbox(self)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.lbx.grid(row=0, column=0, sticky=tk.NSEW)
        self.lbx.bind("<Double-Button-1>", self.lbxdoubleclick)
        self.sb1 = tk.Scrollbar(self, orient=tk.VERTICAL)
        self.sb1.grid(row=0, column=1, sticky=tk.NS + tk.E)
        self.sb1.config(command=self.yview)
        self.lbx.configure(yscrollcommand=self.sb1.set)

        for item in years:
            self.lbx.insert(tk.END, item)

    def yview(self, *args):
        self.lbx.yview(*args)

    def lbxdoubleclick(self, event):
        self.lbxlineselected()

    def lbxlineselected(self):
        selection = self.lbx.curselection()
        if len(selection) > 0:
            self.year = int(self.lbx.get(selection[0]))  # take first line
            self.close()

    def show(self):
        self.wait_window()
        return self.year

    def close(self, event=None):
        """ Close progress window """
        self.parent.focus_set()  # put focus back to the parent window
        self.destroy()  # destroy progress window


class ProgressWindow(simpledialog.Dialog):
    def __init__(self, parent, text):
        """ Init progress window """
        tk.Toplevel.__init__(self, master=parent)
        self.parent = parent
        self.text = text
        self.length = 300

        self.maximum = 100

        self.focus_set()  # set focus on the ProgressWindow

        self.wait_visibility()
        self.grab_set()  # make a modal window, so all events go to the ProgressWindow
        self.transient(self.master)  # show only one window in the task bar
        #
        self.title("Downloading data for {}".format(self.text))
        self.resizable(False, False)  # window is not resizable
        # self.close gets fired when the window is destroyed
        self.protocol("WM_DELETE_WINDOW", self.close)
        # Set proper position over the parent window
        self.geometry("300x30+100+100")

        self.num = tk.IntVar()

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.pgb = ttk.Progressbar(
            self, maximum=self.maximum, orient="horizontal", length=self.length, variable=self.num, mode="determinate"
        )
        self.pgb.grid(row=0, column=0, sticky=tk.NSEW)

        self.num.set(0)

    def set(self, value):
        self.num.set(value)
        self.pgb.update()

    def close(self, event=None):
        """ Close progress window """
        self.master.focus_set()  # put focus back to the parent window
        self.destroy()  # destroy progress window


class SolarviewApp:
    def __init__(self, parent):

        self.parent = parent
        self.parent.title("Solarview - Growatt server annual overview")
        self.readinifile()

        self.createmenubar(self.parent)

        self.prj = Projection()

        self.canvas = tk.Canvas(
            self.parent,
            scrollregion=(0, 0, self.prj.width, self.prj.height),
            width=self.prj.width,  # Ruud: make it not too large
            height=self.prj.height,  # Ruud: make it not too large
            bg="white",
            bd=0,
            highlightthickness=0,
        )

        self.canvas.grid(row=0, column=0)

        self.make_scrollbars()
        self.canvas.update()

        self.year = dt.datetime.now().year

        self.pgw = ProgressWindow(None, str(self.year))
        self.gsd = GrowattServerData(self.year, setprogress=self.pgw.set)
        self.pgw.close()

        self.create_image_pil()
        self.imagetk = ImageTk.PhotoImage(self.image)
        self.canvas.create_image(0, 0, image=self.imagetk, anchor=tk.NW)

        self.canvas.update()

    def readinifile(self):
        config = configparser.ConfigParser()
        if not Path(g.inifilename).exists():
            raise FileNotFoundError(g.inifilename + " not found")

        config.read(g.inifilename)

        g.username = config["ini"]["username"].strip("\"'")
        g.password = config["ini"]["password"].strip("\"'")
        g.pickle_dir = Path(config["ini"]["pickle_dir"].strip("\"'"))
        g.pickle_template = config["ini"]["pickle_template"].strip("\"'")

        if debug:
            print(g.username, g.password, g.pickle_dir, g.pickle_template)

    def make_scrollbars(self):
        sy = tk.Scrollbar(orient=tk.VERTICAL, command=self.canvas.yview)
        sy.grid(row=0, column=1, sticky=tk.NS)
        self.canvas.configure(yscrollcommand=sy.set)

        sx = tk.Scrollbar(orient=tk.HORIZONTAL, command=self.canvas.xview)
        sx.grid(row=1, column=0, sticky=tk.EW)
        self.canvas.configure(xscrollcommand=sx.set)

        top = self.canvas.winfo_toplevel()
        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)
        self.canvas.grid(row=0, column=0, sticky=tk.NSEW)

    def createmenubar(self, root):
        menubar = tk.Menu(root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Select year", command=self.select_year)
        filemenu.add_command(label="Save image", command=self.save_image)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=root.destroy)
        menubar.add_cascade(label="File", menu=filemenu)
        root.config(menu=menubar)

    def draw_grid_pil(self, draw, font, fontbig):
        """
        draw the grid lines
        """
        y_high = self.prj.topmargin - 20
        y_low = self.prj.height - self.prj.bottommargin
        for m in range(1, 14):
            """ determine x-coord of 1st of month """
            if m < 13:  # start of month
                day_diff = isodate_diff("{:4}-{:02}-01".format(self.gsd.year, m), "{:4}-01-01".format(self.gsd.year))
            else:  # start of next year (end line)
                day_diff = isodate_diff("{:4}-01-01".format(self.gsd.year + 1), "{:4}-01-01".format(self.gsd.year))

            x = self.prj.leftmargin + day_diff * (self.prj.pixels_per_day) - (self.prj.pixels_per_day / 2)
            draw.line([(x, y_high), (x, y_low)], fill=0, width=1)

            if m < 13:  # draw month name
                text = calendar.month_name[m]
                bd, hg = draw.textsize(text)
                draw.text((x + 15 * (self.prj.linewidth + 1) - bd / 2, y_low + hg / 2), text=text, fill=0, font=font)

        """ draw hour lines """
        x_min = self.prj.leftmargin
        x_max = self.prj.width - self.prj.rightmargin
        y_max = self.prj.height - self.prj.bottommargin

        """ for 5:00 to 22:00 """
        for h in range(5, 23):
            y = self.prj.fiveoclockbase - (h - 5) * self.prj.pixels_per_hour
            draw.line((x_min, y, x_max, y), fill=(196, 196, 196), width=1)
            text = "{:1}:00".format(h)
            bd, hg = draw.textsize(text)
            draw.text((x_min - bd, y - hg / 2), text=text, fill=(128, 128, 128), font=font)

        """ draw kWh- lines  0 .. 25 """
        for p in range(0, 26, 5):
            y_pos = y_max - p * self.prj.pixels_per_kwh
            draw.line([x_min, y_pos, x_max, y_pos], fill=(196, 196, 196), width=1)
            text = "{:1}".format(p)
            bd, hg = draw.textsize(text)
            draw.text((x_min - 10 - bd / 2, y_pos - hg / 2), text=text, fill=(128, 128, 128), font=font)
        text = "kW"
        bd, hg = draw.textsize(text)
        y_pos = y_max + 10 - 30 * self.prj.pixels_per_kwh
        draw.text((x_min - bd / 2 - 10, y_pos - hg / 2), text=text, fill=(128, 128, 128), font=font)

    def plot_production_pil(self, draw, font):
        """
        Plot production collected from GrowattShinephoneServerdata
        """
        if debug:
            print("Plot production gsd pil")
            print("DEBUG", self.gsd.days.keys())
        for d in self.gsd.days.keys():
            """ x is x-coord of this day"""
            x = (
                self.prj.leftmargin
                + isodate_diff(d, "{:4}-01-01".format(self.gsd.year)) * (self.prj.pixels_per_day)
                + self.prj.pixels_per_day / 2
            )
            """
            Plot day volume, only if detailed day data available
            """
            if len(self.gsd.days[d].samples) > 0:
                pa_today = self.gsd.days[d].todayenergy
                y_low = self.prj.height - self.prj.bottommargin
                draw.line(
                    [x, y_low, x, y_low - (pa_today) * self.prj.pixels_per_kwh],
                    fill=(0, 0, 128),
                    width=self.prj.linewidth,
                )

            """
            Plot heatmapdata
            """
            for ts in sorted(self.gsd.days[d].samples):
                time_of_day_m = isotime_to_m(ts)
                color = self.prj.power_to_color(self.gsd.days[d].samples[ts])
                y = self.prj.fiveoclockbase - ((time_of_day_m - 5 * 60) / 60) * self.prj.pixels_per_hour
                draw.line([x, y, x, y - (5.0 / 60) * self.prj.pixels_per_hour], fill=color, width=self.prj.linewidth)

    def draw_legend_pil(self, draw, font):
        legend_pos = (self.prj.width - self.prj.rightmargin - 140, self.prj.height - self.prj.bottommargin - 220)
        lineno = 0
        power_legend = self.prj.power_legend()
        for l in power_legend:
            colorstr = l[0]
            textstr = l[1]
            draw.line(
                [legend_pos[0], legend_pos[1] - lineno * 10, legend_pos[0] + 40, legend_pos[1] - lineno * 10],
                fill=colorstr,
                width=8,
            )
            draw.text(
                [legend_pos[0] + 45, legend_pos[1] - 5 - lineno * 10],
                text=textstr,
                fill=(0, 0, 0),
                font=font,
                align="left",
            )
            lineno += 1

    """ Plot location and power generated this year"""

    def plot_title_pil(self, draw, font, bigfont):
        title1str = "{:4}".format(self.gsd.year)
        bd, hg = draw.textsize(title1str)
        draw.text(
            (self.prj.leftmargin + bd / 2, self.prj.topmargin - hg / 2 - 10), text=title1str, fill=0, font=bigfont
        )

        title2str = self.gsd.plant_id + " " + self.gsd.plant_name
        bd, hg = draw.textsize(title2str)
        title2pos = (self.prj.leftmargin + self.prj.width / 2 - bd / 2, self.prj.topmargin - hg / 2 - 10)
        draw.text(title2pos, text=title2str, fill=(0, 0, 0), font=font, align="left")

        title3str = "{:0.0f} kWh".format(self.gsd.yearproduction)
        bd, hg = draw.textsize(title3str)
        title3pos = (
            self.prj.width - self.prj.rightmargin - self.prj.leftmargin - bd + 10,
            self.prj.topmargin - hg / 2 - 10,
        )
        draw.text(title3pos, text=title3str, fill=(0, 0, 0), font=bigfont)

    def create_image_pil(self):

        self.image = Image.new("RGB", (self.prj.width, self.prj.height), (255, 255, 255))  # white
        try:
            font = ImageFont.truetype("arial.ttf", 10)
            fontbig = ImageFont.truetype("arial.ttf", 18)
        except IOError:
            try:
                font = ImageFont.truetype("LiberationSans-Regular.ttf", 10)
                fontbig = ImageFont.truetype("LiberationSans-Regular", 18)
            except IOError:
                font = ImageFont.load_default()
                fontbig = ImageFont.load_default()

        idraw = ImageDraw.Draw(self.image)

        self.draw_grid_pil(idraw, font, fontbig)
        self.plot_production_pil(idraw, font)
        self.draw_legend_pil(idraw, font)
        self.plot_title_pil(idraw, font, fontbig)

    def select_year(self):
        """ Open modal window """
        selyear = YearSelector(self.parent, self.gsd.yearsavailable).show()
        if debug:
            print("Selected year: {}".format(selyear))

        if selyear > 0:
            self.year = selyear

            self.pgw = ProgressWindow(None, str(self.year))
            self.gsd = GrowattServerData(self.year, setprogress=self.pgw.set)
            self.pgw.close()

            self.create_image_pil()
            self.imagetk = ImageTk.PhotoImage(self.image)
            self.canvas.create_image(0, 0, image=self.imagetk, anchor=tk.NW)

            self.canvas.update()

    def save_image(self):
        myFormats = [("Portable Network Graphics", "*.png"), ("JPEG / JFIF", "*.jpg")]
        filename = filedialog.asksaveasfilename(filetypes=myFormats)
        if filename:
            file = Path(filename)
            if not file.suffix:
                file = file.with_suffix(".png")
            self.image.save(file)

    def donothing(self):
        pass


def main():

    mainwindow = tk.Tk()
    SolarviewApp(mainwindow)
    mainwindow.mainloop()


if __name__ == "__main__":
    main()
