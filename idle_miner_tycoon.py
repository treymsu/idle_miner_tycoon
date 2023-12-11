"""Play Idle Miner Tycoon"""
# pip install winsdk pyautogui screen_ocr[winrt] wheel pywin32 opencv-python
import argparse
import time
import os
import logging
import sys
from collections import namedtuple
from enum import Enum
import win32gui
import pyautogui
import pyscreeze
import screen_ocr

DEBUG = True
LOW_RESOLUTION = (423, 726)
HIGH_RESOLUTION = (659, 1131)
OCR_READER = screen_ocr.Reader.create_quality_reader()
SCALE = 1
BLUESTACKS = None
SCRIPT_DIR = os.path.dirname(__file__)

os.chdir(SCRIPT_DIR)
logger = logging.getLogger(__name__)
formatter = logging.Formatter(
    "%(asctime)s:%(levelname)s:%(lineno)d:%(message)s", datefmt="%H:%M:%S")
logger.setLevel(logging.DEBUG if DEBUG else logging.INFO)
ch = logging.StreamHandler()
ch.setFormatter(formatter)
logger.addHandler(ch)

Point = namedtuple("Point", ["x", "y"])
Box = namedtuple("Box", ["x", "y", "w", "h"])
BoundingBox = namedtuple("BoundingBox", ["left", "top", "right", "bottom"])
Window = Enum("Window", ["SHAFT", "MANAGER_CHOOSER", "LEVEL_UP", "MINE_OVERVIEW"])
MineArea = Enum("MineArea", ["MINESHAFT", "ELEVATOR", "WAREHOUSE"])
MineMode = Enum("MineMode", ["EVENT", "MAINLAND", "FRONTIER", "REGULAR"])


def window_callback(hwnd, extra=None):
    """Use win32gui to find the location of the bluestacks window"""
    global BLUESTACKS, SCALE
    title = win32gui.GetWindowText(hwnd)
    if "bluestacks app player" not in title.lower():
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w, h = right - left, bottom - top
    if BLUESTACKS == (left, top, w, h):
        return  # Nothing changed
    logger.info("'%s', Loc: (%d, %d), Size: (%d, %d)", title, left, top, w, h)
    if (w, h) not in (HIGH_RESOLUTION, LOW_RESOLUTION):
        logger.error("Resolution not found: (%d, %d)", w, h)
        w, h = HIGH_RESOLUTION
    win32gui.MoveWindow(hwnd, 0, 0, w, h, True)
    if (w, h) == HIGH_RESOLUTION:
        logger.info("Changing to high-resolution mode")
        os.chdir(os.path.join(SCRIPT_DIR, f"{w}x{h}"))
        SCALE = 1.557851
    elif (w, h) == LOW_RESOLUTION:
        logger.info("Changing to low-resolution mode")
        os.chdir(os.path.join(SCRIPT_DIR, f"{w}x{h}"))
    else:
        logger.error("Bad resolution")
        sys.exit(-1)
    BLUESTACKS = Box(left, top, w, h)


win32gui.EnumWindows(window_callback, None)
if BLUESTACKS is None:
    logger.error("Failed to find bluestacks, exiting")
    sys.exit(-1)


class Region:
    """Areas relative to the top left of the Bluestacks window"""

    def __init__(self, left, top, right, bottom, anchor=None, debug=False):
        self.anchor = anchor if anchor is not None else BLUESTACKS
        # Original unchanging values
        self._left = round(left * SCALE)
        self._top = round(top * SCALE)
        self._right = round(right * SCALE)
        self._bottom = round(bottom * SCALE)
        self.width = self._right - self._left
        self.height = self._bottom - self._top
        self.left = self._left
        self.top = self._top
        self.right = self._right
        self.bottom = self._bottom
        if debug:
            self.draw()

    def _update(self):
        """Update all the members based on bluestacks coords"""
        offset_x = self.anchor.x
        offset_y = self.anchor.y
        self.left = self._left + offset_x
        self.right = self._right + offset_x
        self.top = self._top + offset_y
        self.bottom = self._bottom + offset_y

    def box(self):
        """Return a Box(x, y, w, h) for image location"""
        self._update()
        return Box(self.left, self.top, self.width, self.height)

    def bounding_box(self):
        """Return a BoundingBox(l, t, r, b) for OCR"""
        self._update()
        return BoundingBox(self.left, self.top, self.right, self.bottom)

    def ocr(self):
        """Read text inside region (lowercase, no spaces, no periods)"""
        if DEBUG:
            self.draw()
        bbox = self.bounding_box()
        assert bbox.bottom - bbox.top >= 20
        results = OCR_READER.read_screen(bbox).as_string()
        results = results.strip().replace(" ", "").replace(".", "").lower()
        return results

    def draw(self, duration=0.5, times=1):
        """Move mouse around border of region"""
        self._update()
        for _ in range(times):
            if (self.left, self.top) != (0, 0):
                pyautogui.moveTo(self.left, self.top, duration=0)
            pyautogui.moveTo(self.right, self.top, duration=duration)
            pyautogui.moveTo(self.right, self.bottom, duration=duration)
            pyautogui.moveTo(self.left, self.bottom, duration=duration)
            if (self.left, self.top) != (0, 0):
                pyautogui.moveTo(self.left, self.top, duration=duration)


class Loc:
    """A Point coordinate"""

    def __init__(self, x, y, anchor=None, debug=False):
        """Create a new coordinate"""
        self.debug = debug or DEBUG
        self.x = x
        self.y = y
        self.anchor = anchor if anchor is not None else BLUESTACKS

    def _update(self):
        """Update coords based on window position"""
        rel_x = round(self.x * SCALE) + self.anchor.x
        rel_y = round(self.y * SCALE) + self.anchor.y
        return Point(int(rel_x), int(rel_y))

    def loc(self):
        """Get the coords"""
        return self._update()

    def click(self):
        """Click the coords"""
        self.draw()
        pyautogui.click(self._update())

    def get_color(self):
        """Get the pixel color"""
        x, y = self._update()
        self.draw()
        return pyautogui.pixel(x, y)

    def draw(self, wait=0.5):
        """Put pointer on coords"""
        if not self.debug:
            return
        pyautogui.moveTo(self._update())
        time.sleep(wait)


class Color:
    """Class to determine if a pixel is within a color range"""

    def __init__(self, color, crange=(0, 0, 0)):
        self.color = color
        self.crange = crange

    def __eq__(self, in_color):
        min_color = (
            self.color[0] - self.crange[0],
            self.color[1] - self.crange[1],
            self.color[2] - self.crange[2],
        )
        max_color = (
            self.color[0] + self.crange[0],
            self.color[1] + self.crange[1],
            self.color[2] + self.crange[2],
        )
        if (
            min_color[0] <= in_color[0] <= max_color[0]
            and min_color[1] <= in_color[1] <= max_color[1]
            and min_color[2] <= in_color[2] <= max_color[2]
        ):
            return True
        return False


class IdleMinerTycoon:
    """Play the game"""

    def __init__(self):
        self.confidence = 0.8
        self.last_edgar_time = time.perf_counter()
        self.last_edgar_search_time = time.perf_counter()
        self.last_upgrade_time = 0
        self.area_needs_leveling: MineArea = MineArea.MINESHAFT
        self.mine: MineMode = MineMode.REGULAR
        self.maxed_elevator = False
        self.maxed_warehouse = False
        self.maxed_mineshaft = False
        self.current_mgr = {
            MineArea.MINESHAFT: {"known": False, "boosted": True},
            MineArea.ELEVATOR: {"known": False, "boosted": True},
            MineArea.WAREHOUSE: {"known": False, "boosted": True},
        }
        self.region_game = Region(0, 32, BLUESTACKS.w - 32, BLUESTACKS.h)
        self.always_buttons = [
            "free.png", "edgar.png", "free-idle.png",  # "30m-skip.png",
            "remove-barrier.png", "collect.png", "free-idle.png", "free.png",
            "x.png", "x2.png", "x3.png", "x4.png", "claim.png",
            "close-blue.png", "get.png",
        ]
        self.mgrs = {
            # Mineshaft
            'pebble': 2, 'mrturner': 0.5, 'rangersue': 1, 'zigalvani': 1,
            'blingsley': 1, 'chester': 5, 'goodmanjr': 5, 'gordon': 5,
            'greenidler': 1, 'cliffwalker': 0.5, 'drsteiner': 5,
            'rabbidblingsley': 2.5, 'sirlorenzo': 1,
            # Elevator
            'queenaurora': 2.5, 'drlilly': 5, 'damianjones': 5, 'sojo': 5,
            'mrsgoodman': 5, 'leevatori': 5, 'ezioauditore': 1,
            'zephyria': 2.5,
            # Warehouse
            'professormaple': 1, 'drnova': 5, 'luxario': 1, 'mark': 5,
            'mrgoodman': 5, 'octaviadevere': 2.5, 'chriscapella': 5,
            'jadekim': 5,
        }
        # The heading at the top of the panel to choose managers.
        # Mineshaft XX Manager, Elevator Manager, Warehouse
        self.region_manager_chooser_heading = Region(88, 105, 304, 132)
        # next_change_time[area] = time.perf_counter()
        self.next_change_time = {
            MineArea.MINESHAFT: time.perf_counter(),
            MineArea.ELEVATOR: time.perf_counter(),
            MineArea.WAREHOUSE: time.perf_counter()
        }
        self.colors = {
            # Boost colors
            "mgr_red": Color((255, 94, 71), (0, 4, 1)),  # Super Managers
            "mgr_green": Color((205, 255, 132), (6, 0, 3)),  # Executive
            "mgr_yellow": Color((254, 254, 180), (1, 1, 3)),  # Senior
            "mgr_blue": Color((198, 250, 252), (4, 2, 1)),  # Junior
            # Mine overview colors
            "mo_orange": Color((254, 210, 2), (1, 30, 0)),
            # Normal orange for boost icon (177, 113, 1)
            "cycle_orange": Color((177, 113, 5), (10, 10, 6)),
            # Bright orange, when mgr is already assigned in this mine (255, 210, 64)
            "cycle_orange_dark": Color((252, 210, 64), (10, 10, 10)),
            "upgrade_arrow_left": Color((255, 230, 123), (5, 10, 10)),
            "upgrade_arrow_right": Color((255, 208, 2), (0, 10, 10)),
        }

    def find_image_timeout(self, image, timeout, click=False, confidence=None, region=None):
        """Wrap the find_image function with a timeout to search for N seconds"""
        end_time = time.perf_counter() + timeout
        while time.perf_counter() < end_time:
            ret = self.find_image(image, click, confidence, region)
            if ret:
                return ret
        logger.warning("Couldn't find %s after %ds", image, timeout)
        return None

    def find_image(self, image, click=False, confidence=None, region=None):
        """Search for an image in a region"""
        c = confidence or self.confidence
        r = region or self.region_game
        action = "Clicked" if click else "Found"
        if not os.path.exists(image):
            logger.error("%s does not exist", image)
            return None
        try:
            loc = pyautogui.locateCenterOnScreen(image, confidence=c, region=r.box())
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            return None
        if loc is not None:
            if click:
                pyautogui.click(loc)
            logger.debug("%s %s at %s", action, image, loc)
        return loc

    def locate_center(self, image, confidence=None, region=None):
        """Wrapper to catch exceptions"""
        if not os.path.exists(image):
            logger.error("%s does not exist", image)
            return False
        r = region or self.region_game
        c = confidence or self.confidence
        try:
            result = pyautogui.locateCenterOnScreen(
                image, confidence=c, region=r.box())
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            return None
        return result

    def locate_all(self, image, confidence=None, region=None):
        """Wrapper to catch exceptions"""
        if not os.path.exists(image):
            logger.error("%s does not exist", image)
            return False
        r = region or self.region_game
        c = confidence or self.confidence
        try:
            result = list(pyautogui.locateAllOnScreen(
                image, confidence=c, region=r.box()))
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            return None
        return result

    def verify_in_shaft(self, timeout=3):
        """Make sure it looks like we're in a mineshaft"""
        end_time = time.perf_counter() + timeout
        while time.perf_counter() < end_time:
            all_levels = self.locate_all("level.png", confidence=0.7)
            if all_levels:
                return True
        logger.error("Can't find shaft!")
        return True

    def verify_in_game(self, timeout=3):
        """Make sure it looks like we're still in the game"""
        end_time = time.perf_counter() + timeout
        iters = 0
        while time.perf_counter() < end_time:
            if self.find_image("shovel.png"):
                return True
            if self.find_image("shop.png"):  # get shop with !
                return True
            if self.find_image("frontier-shop.png"):
                return True
            if self.find_image("frontier-shop2.png"):
                return True
            iters += 1
        logger.error("Can't find shovel/shop to verify in game! %d", iters)
        return False

    def verify_in_manager_window(self, area: MineArea):
        """Make sure the manager choosing window is in the foreground"""
        region = self.region_manager_chooser_heading
        end_time = time.perf_counter() + 3
        iters = 0
        while time.perf_counter() < end_time:
            ocr = region.ocr()
            logger.debug("OCR: Manager window title: %s", ocr)
            if "manager" in ocr or "hanager" in ocr:
                return True
            elif (area == MineArea.MINESHAFT and ("mine" in ocr or "shaft" in ocr)):
                return True
            elif (area == MineArea.WAREHOUSE and ("ware" in ocr or "house" in ocr)):
                return True
            else:
                if "ele" in ocr or "vator" in ocr:
                    return True
            time.sleep(0.25)
            iters += 1
        logger.error("Manager window not found via OCR: %s (%d iters)", ocr, iters)
        return False

    def verify_in_mine_overview(self):
        """Make sure we're in the Mine Overview window"""
        heading = Region(95, 95, 275, 125)
        end_time = time.perf_counter() + 3
        while time.perf_counter() < end_time:
            ocr = heading.ocr()
            logger.debug("Mine overview OCR = %s", ocr)
            if "mineoverview" in ocr or "måneovervåew" in ocr or "over" in ocr:
                logger.debug("Mine overview found via OCR")
                return True
            time.sleep(0.25)
        logger.error('Mine overview not found in "%s"', ocr)
        return False

    def goto_mineshaft_top(self, top=True):
        """Go to the top of the mineshaft"""
        if not self.verify_in_shaft():
            return False
        if top:
            up_arrow = Loc(25, 603)
            up_arrow.click()
        else:
            down_arrow = Loc(25, 655)
            down_arrow.click()
        time.sleep(2)
        return True

    def goto_mineshaft_bottom(self):
        """Go to the bottom of the mineshaft"""
        return self.goto_mineshaft_top(top=False)

    def find_last_manager(self):
        """Get the location of the last manager in the mineshaft"""
        loc = self.get_last_level()
        if not loc:
            logger.error("No last manager found")
            return None
        return Point(loc[0] - round(150 * SCALE), loc[1])

    def get_last_level(self):
        """Get the location of the bottom-most 'Level' button"""
        self.goto_mineshaft_bottom()
        time.sleep(2)  # extra time
        all_levels = self.locate_all("level.png", confidence=0.7)
        if not all_levels:
            logger.error("No level icon found")
            return None
        return sorted(all_levels, key=lambda x: x[1])[-1]

    def goto_mine_overview_top(self, top=True):
        """Go to the top of the mine overview window"""
        if top:
            arrow = Loc(50, 563)
        else:
            arrow = Loc(50, 606)
        arrow.click()
        time.sleep(2)
        return True

    def goto_mine_overview_bottom(self):
        """Go to the bottom of the mine overview window"""
        return self.goto_mine_overview_top(top=False)

    def open_mine_overview(self):
        """Open the Mine Overview window"""
        if not self.find_image("shovel.png", click=True):
            logger.error("Shovel not found to open Mine Overview")
            return False
        return True

    def mine_overview(self):
        """Open mine overview, see what area needs leveling up"""
        if not self.open_mine_overview():
            logger.error("Couldn't open mine overview")
            return False

        if not self.verify_in_mine_overview():
            logger.error("Can't verify in mine overview")
            return False

        self.goto_mine_overview_top()

        loc = self.find_image_timeout("exclamation.png", timeout=2)
        if loc is None:
            logger.error("Exclamation not found")
            time.sleep(1)
            pyautogui.press(["esc"])
            time.sleep(1)
            return False

        prev_anl = self.area_needs_leveling
        y = loc[1] - self.region_game.top
        if y < round(200 * SCALE):
            self.area_needs_leveling = MineArea.WAREHOUSE
            if self.maxed_warehouse:
                if self.maxed_elevator:
                    self.area_needs_leveling = MineArea.MINESHAFT
                else:
                    self.area_needs_leveling = MineArea.ELEVATOR
        elif y < round(350 * SCALE):
            self.area_needs_leveling = MineArea.ELEVATOR
            if self.maxed_elevator:
                if self.maxed_warehouse:
                    self.area_needs_leveling = MineArea.MINESHAFT
                else:
                    self.area_needs_leveling = MineArea.WAREHOUSE
        else:
            self.area_needs_leveling = MineArea.MINESHAFT
            if self.maxed_mineshaft:
                if self.maxed_elevator:
                    self.area_needs_leveling = MineArea.WAREHOUSE
                else:
                    self.area_needs_leveling = MineArea.ELEVATOR

        if self.area_needs_leveling != prev_anl:
            logger.info("%s area needs leveling!", self.area_needs_leveling.name)

        active_boost_color = Color((255, 243, 115))
        recovery_color = Color((183, 183, 183))
        background_blue_color = Color((11, 92, 147))  # when ready to boost
        bars = dict()
        bars[MineArea.WAREHOUSE] = Loc(174, 273)
        bars[MineArea.ELEVATOR] = Loc(174, 417)
        bars[MineArea.MINESHAFT] = Loc(174, 613)
        for area in MineArea:
            if self.mine in (MineMode.FRONTIER, ):
                break
            if not self.current_mgr[area]["known"]:
                if area == MineArea.MINESHAFT:
                    self.goto_mine_overview_bottom()
                else:
                    self.goto_mine_overview_top()

                loc = bars[area]
                pix = loc.get_color()
                if pix == active_boost_color:
                    if self.current_mgr[area]["boosted"]:
                        logger.debug("%s boost in progress", area.name)
                    else:
                        logger.info("%s manager is boosted", area.name)
                    self.current_mgr[area]["boosted"] = True
                else:
                    self.current_mgr[area]["boosted"] = False
                    if pix == recovery_color:
                        logger.debug("%s manager recovering", area.name)
                    elif pix == background_blue_color:
                        logger.debug("%s manager ready to boost", area.name)
                    else:
                        logger.debug("%s manager ready to boost %s", area.name, pix)

        # Exit mine overview
        pyautogui.press(["esc"])
        time.sleep(1)
        return True

    def max_all_mines(self):
        """Starting from the last mine, max all of them"""
        left_arrow = Loc(10, 380)
        heading_region = Region(41, 94, 317, 131)
        for _ in range(35):
            left_arrow.click()
            time.sleep(1)
            text = heading_region.ocr()  # This OCR is very unreliable
            if not (self.find_image("mineshaft.png") or 'mineshaft' in text
                    or 'mdne' in text or 'mbne' in text):
                logger.error("Not looking at a mineshaft, not maxing: %s", text)
                return

            # We're in a non-maxed shaft, let's try to click upgrade
            if self.find_image("upgrade.png", click=True):
                logger.info("Maxing mine '%s'", text)
            else:
                logger.info("%s already maxed", text)

            # Mineshaft 1 Level 800
            if "ft1lev" in text:
                logger.info("Done maxing everything!")
                return
            time.sleep(0.5)
            # We might have maxed first 5 only...
            # maxed_mineshafts = True

    def level_up(self):
        """Upgrade an area"""
        self.mine_overview()
        logger.debug("Leveling up %s", self.area_needs_leveling.name)
        if not self.verify_in_game():
            logger.error("Not in game, canceling level up")
            return
        if not self.verify_in_shaft():
            logger.error("Not in shaft, canceling level up")
            return

        if self.area_needs_leveling == MineArea.MINESHAFT:
            self.goto_mineshaft_bottom()
            loc = self.get_last_level()
            if loc is None:
                logger.error("Can't find last level")
                return False
            arrow_loc = loc[0], loc[1] - 5
        elif self.area_needs_leveling == MineArea.ELEVATOR:
            self.goto_mineshaft_top()
            loc = Loc(61, 302).loc()  # Elevator "Level" button
            arrow_loc = Loc(65, 282).loc()
        elif self.area_needs_leveling == MineArea.WAREHOUSE:
            self.goto_mineshaft_top()
            loc = Loc(345, 302).loc()  # Warehouse "Level" button
            arrow_loc = Loc(325, 282).loc()
        else:
            logger.error("Unknown area: %s", self.area_needs_leveling)
            return False

        # Look for arrow
        upgradable = False
        for i in range(0, 10, 3):
            x, y = arrow_loc[0], arrow_loc[1] - i
            if DEBUG:
                pyautogui.moveTo(x, y)
            pix = pyautogui.pixel(int(x), int(y))
            if (self.colors["upgrade_arrow_left"] == pix
                    or self.colors["upgrade_arrow_right"] == pix):
                logger.debug("Found upgrade arrow")
                upgradable = True
                break
        if not upgradable:
            logger.debug("Upgrade arrow not found")
            return False
        if DEBUG:
            pyautogui.moveTo(loc)
        pyautogui.click(loc)
        time.sleep(1)
        self.find_image("max-selected.png", click=True)
        self.find_image("max-unselected.png", click=True)
        time.sleep(0.5)
        if self.find_image("upgrade.png", click=True):
            self.last_upgrade_time = time.perf_counter()
            time.sleep(0.5)

        if self.find_image("maxed-upgrades.png"):
            if self.area_needs_leveling == MineArea.MINESHAFT:
                self.max_all_mines()
            elif self.area_needs_leveling == MineArea.ELEVATOR:
                self.maxed_elevator = True
            elif self.area_needs_leveling == MineArea.WAREHOUSE:
                self.maxed_warehouse = True

        pyautogui.press(["esc"])
        time.sleep(1)

    def _find_next_mgr(self, area, mgr_name=None, boost=True):
        assign_buttons = self.locate_all("assign.png", confidence=0.7)
        if not assign_buttons:
            logger.error("No assign buttons found")
            return False

        logger.debug("%u assign buttons: %s", len(assign_buttons), assign_buttons)
        for assign_button in assign_buttons:
            found_it = False
            if DEBUG:
                pyautogui.moveTo(assign_button)
            if mgr_name is None:
                # Find next SM that's ready
                for xx in range(0, 24, 4):
                    for yy in range(0, 24, 4):
                        x_offset = round(30 * SCALE)
                        y_offset = round(33 * SCALE)
                        x = int(assign_button[0] + x_offset + xx)
                        y = int(assign_button[1] + y_offset + yy)
                        if DEBUG:
                            # This takes a really long time
                            pyautogui.moveTo(x, y)
                        pix = pyautogui.pixel(x, y)
                        if (
                            self.colors["cycle_orange"] == pix
                            or self.colors["cycle_orange_dark"] == pix
                        ):
                            found_it = True
                            break
                    if found_it:
                        break
                if not found_it:
                    logger.debug("Doesn't look ready, moving to next mgr: %s", pix)
                    continue
                logger.debug("Looks ready, assigning mgr. pix=%s", pix)
            else:
                # Find manager by name
                pt = Point(assign_button.left, assign_button.top)
                mgr_name_region = Region(-120, -12, -20, 12, pt)
                text = mgr_name_region.ocr()
                logger.info("Manager name next to assign button: %s", text)
                if mgr_name != text:
                    logger.debug("Manager %s isn't %s", text, mgr_name)
                    continue
            # Assign manager
            if DEBUG:
                pyautogui.moveTo(assign_button[0], assign_button[1])
            pyautogui.click(assign_button[0], assign_button[1])
            self.current_mgr[area]["known"] = False
            self.current_mgr[area]["boosted"] = False
            time.sleep(1)
            if (self.find_image("assign-anyway.png", click=True)
                    or self.find_image("assign-anyway2.png", click=True)):
                time.sleep(2)
            # Find manager name
            name_region = Region(145, 178, 242, 198)
            mgr_name = name_region.ocr()
            if mgr_name in self.mgrs:
                self.current_mgr[area]["known"] = True
                active_time = self.mgrs[mgr_name] * 60  # convert min to sec
            else:
                active_time = 30  # minimum time
            # Boost manager
            if boost:
                unassign_loc = self.find_image("unassign.png")
                if unassign_loc is None:
                    logger.warning("Couldn't find unassign button")
                    continue
                boost_loc_x = unassign_loc[0] + round(20 * SCALE)
                boost_loc_y = unassign_loc[1] + round(35 * SCALE)
                boost_loc = (boost_loc_x, boost_loc_y)
                if DEBUG:
                    pyautogui.moveTo(boost_loc)
                logger.info("Boosting manager %s for %ds", mgr_name, active_time)
                self.current_mgr[area]["boosted"] = True
                pyautogui.click(boost_loc)
                self.next_change_time[area] = time.perf_counter() + active_time
            return True
        logger.debug("No boostable managers found, need to scroll")
        return False

    def open_manager_window(self, area):
        """Open the manager window (by clicking on a manager in an area)"""
        if area == MineArea.MINESHAFT:
            self.goto_mineshaft_bottom()
            mgr_loc = self.find_last_manager()
            if not mgr_loc:
                logger.warning("Couldn't find last manager")
                return False
            pyautogui.click(mgr_loc)
        elif area == MineArea.ELEVATOR:
            self.goto_mineshaft_top()
            mgr_loc = Loc(53, 405)
            mgr_loc.click()
        else:
            self.goto_mineshaft_top()
            mgr_loc = Loc(340, 405)
            mgr_loc.click()
        return self.verify_in_manager_window(area)

    def _cycle_managers(self, area, mgr_name=None, boost=True):
        """Change the super manager for a particular area"""
        logger.debug("Cycling manager in %s area", area)
        if not self.verify_in_shaft():
            return

        self.open_manager_window(area)
        # pick super manager tab
        super_manager_tabs = [
            "super-managers-tab.png",
            "super-managers-tab-dark.png",
            "super-managers-tab-dark2.png",
        ]
        for img in super_manager_tabs:
            loc = self.find_image(img)
            if loc is not None:
                pyautogui.click(loc)
                break

        # Find next manager with orange below assign button
        scrolls = 5
        if area in (MineArea.WAREHOUSE, MineArea.ELEVATOR):
            scrolls = 4
        boosted = False
        for i in range(scrolls):
            if not self.verify_in_manager_window(MineArea.MINESHAFT):
                logger.error("Not in Mineshaft Manager window")
                break

            if self._find_next_mgr(area, mgr_name=mgr_name, boost=boost):
                boosted = True
                break

            # Scroll
            if i == scrolls - 1:
                # Don't scroll on the last iteration
                break
            drag_start = (
                self.region_game.left + self.region_game.right // 2,
                self.region_game.top + round(500 * SCALE),
            )
            # 115 is about 1 manager size chunk
            pyautogui.moveTo(drag_start)
            time.sleep(0.5)
            pyautogui.dragRel(xOffset=0, yOffset=(-200 * SCALE), duration=2)
            time.sleep(3)
            logger.debug("trying again...")
        if not boosted:
            logger.info("No boostable manager found, waiting 2 min")
            self.next_change_time[area] = time.perf_counter() + 2*60
        pyautogui.press(["esc"])
        time.sleep(1)
        return

    def cycle_managers(self):
        """Change managers in all areas"""
        if self.mine not in (MineMode.REGULAR, MineMode.EVENT):
            logger.debug("Not cycling managers for %s mode", self.mine.name)
            return False
        for area in (MineArea.MINESHAFT, MineArea.ELEVATOR, MineArea.WAREHOUSE):
            now = time.perf_counter()
            next_time = self.next_change_time[area]
            delta = abs(now - next_time)
            if now < next_time:
                logger.debug("%s not ready yet %ds", area.name, delta)
                continue
            logger.debug("Past %s change time by %ds", area.name, delta)

            if not self.current_mgr[area]["known"] and self.current_mgr[area]["boosted"]:
                logger.warning("Unknown manager in %s is still boosted", area.name)
                continue
            self._cycle_managers(area, boost=True)
            self.edgar()

    def edgar(self):
        """Search for edgar and click"""
        logger.debug("Searching for Edgar")
        # edgar_region = Region(355, 900, 560, 1060)
        edgar_region = Region(228, 578, 359, 680)
        now = time.perf_counter()
        search_delta = self.last_edgar_search_time - now
        if search_delta > 5:
            logger.info("took %u seconds between edgar searches", search_delta)
        self.last_edgar_search_time = now

        e = self.find_image("edgar.png", click=True, region=edgar_region)
        e2 = self.find_image("edgar-extravaganza.png", click=True, region=edgar_region)
        if e or e2:
            self.last_edgar_time = time.perf_counter()
            # If we clicked edgar, wait 5 seconds for free button
            end_time = time.perf_counter() + 5
            while time.perf_counter() < end_time:
                if self.find_image("free.png", click=True):
                    logger.info("Found edgar")
                    time.sleep(5)
                    break
                time.sleep(0.25)

        now = time.perf_counter()
        if now > self.last_edgar_time + 30 * 60:
            minutes_since = (now - self.last_edgar_time) // 60
            logger.warning("Haven't seen edgar in %d minutes", minutes_since)
            # relaunch_game()

    def new_shaft(self):
        """Open the next shaft, if it's ready"""
        logger.debug("New Shaft")
        new_shaft_region = Region(240, 80, 385, 695)
        for _ in range(5):
            self.goto_mineshaft_bottom()
            loc = self.locate_center("new-shaft.png", region=new_shaft_region)
            if not loc:
                break

            # double check with some color matching
            new_shaft_blue = Color((87, 170, 227), (5, 5, 5))

            # might need to check a few pixels
            check_pixel = int(loc[0] - 5), int(loc[1] - 5)
            pix = pyautogui.pixel(check_pixel[0], check_pixel[1])
            if DEBUG:
                pyautogui.moveTo(check_pixel)

            if pix != new_shaft_blue:
                logger.debug("New shaft button isn't the right color: %s", pix)
                return

            logger.info("Opening new shaft")
            pyautogui.click(loc)
            time.sleep(3)
            self.hire_last_manager()

    def hire_last_manager(self):
        """Hire someone in the last mineshaft mgr spot"""
        logger.debug("Hiring last manager")
        loc = self.find_last_manager()
        if not loc:
            return
        pyautogui.click(loc)
        time.sleep(2)

        # It is ok if this doesn't work
        loc = self.locate_center("dollar-mgr-tab.png")
        if loc:
            pyautogui.click(loc)
            time.sleep(0.5)

        loc = self.locate_center("hire-manager-button.png")
        if not loc:
            logger.warning("Can't find hire-manage-button, trying alternate")
            loc = self.locate_center("hire-manager-button2.png")
        if not loc:
            logger.error("Can't find hire button")
        pyautogui.click(loc)
        logger.info("Hired manager")
        time.sleep(0.5)
        pyautogui.press(["esc"])

    def unlock_barrier(self):
        """Unlock any barrier that can be unlocked"""
        logger.debug("Trying to unlock barrier")
        self.goto_mineshaft_bottom()
        loc = self.locate_center("remove-barrier.png")
        if loc:
            pyautogui.click(loc[0] - 5, loc[1] - 5)
            logger.info("Unlocking barrier!")
            time.sleep(1)

        # Skip timers, too
        self.find_image("skip-no-time.png", click=True)

    def discover_location(self):
        """Take a guess at what kind of mine we're currently in"""
        if not self.verify_in_shaft(timeout=120):
            logger.error("Timed out waiting for us to be in a mineshaft")
        self.goto_mineshaft_top()
        if self.find_image("event-mine.png"):
            mine = MineMode.EVENT
        elif self.find_image("mainland-menu.png"):
            mine = MineMode.MAINLAND
        elif self.find_image("frontier-shop.png") or self.find_image("frontier-shop2.png"):
            mine = MineMode.FRONTIER
        else:
            mine = MineMode.REGULAR
        logger.info("In %s mine.", mine.name)
        self.mine = mine
        return mine

    def close_popups(self):
        """Try to close as many popups as we can"""
        for _ in range(5):
            if self.verify_in_shaft():
                return
            for img in [
                "free.png", "free-idle.png", "skip-no-time.png",
                "collect.png", "x.png", "x2.png", "x3.png", "x4.png",
                "red-x.png", "cancel.png",
            ]:
                if self.find_image(img, click=True):
                    time.sleep(2)
            pyautogui.press(["esc"])
            time.sleep(2)
            self.find_image_timeout("cancel.png", click=True, timeout=5)

    def start_game(self):
        """Open the game from the bluestacks app menu"""
        if self.find_image("idle-miner.png", click=True):
            time.sleep(30)
            self.close_popups()
            self.discover_location()

    def play(self):
        """Run the game"""
        iteration = 1
        self.discover_location()
        while True:
            # Every N cycles, make sure we're still in the expected mine
            # if iteration % 50 == 0:
                # win32gui.EnumWindows(window_callback, None)
                # self.discover_location()
            self.start_game()
            for img in self.always_buttons + ['cancel.png', 'red-x.png']:
                self.find_image(img, click=True)
            self.edgar()
            self.level_up()
            self.edgar()
            self.new_shaft()
            self.unlock_barrier()
            self.edgar()
            self.cycle_managers()
            self.edgar()
            iteration += 1

    def test(self):
        """Function to test a single feature when run with --test"""
        global DEBUG
        DEBUG = True
        # self.mine_overview()
        # self.discover_location()
        # self._cycle_managers(MineArea.ELEVATOR)
        # self._cycle_managers(MineArea.WAREHOUSE)
        # self._cycle_managers(MineArea.MINESHAFT)
        # self.unlock_barrier()
        # self.new_shaft()
        # print("verify_in_game: %s" % self.verify_in_game())
        # for area in (MineArea.ELEVATOR, MineArea.WAREHOUSE, MineArea.MINESHAFT):
        #     self.open_manager_window(area)
        #     print("verify_in_manager_window: %s" % self.verify_in_manager_window(area))
        #     pyautogui.press(["esc"])
        #     time.sleep(0.5)
        # print(f"open_mine_overview: {self.open_mine_overview()})
        # print(f"verify_in_mine_overview: {self.verify_in_mine_overview()}")
        print(f"verify_in_shaft: {self.verify_in_shaft()}")


def play():
    """Entry point"""
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action='store_true')
    args = parser.parse_args()
    imt = IdleMinerTycoon()
    if args.test:
        imt.test()
    else:
        imt.play()


if __name__ == "__main__":
    play()
