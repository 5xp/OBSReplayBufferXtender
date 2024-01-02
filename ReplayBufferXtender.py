import os

import obspython as o
import psutil
import win32api
import win32gui as wgui
import win32process


class ReplayBufferXtender:
    """
    All main script logic lives here
    """

    base_dir = None
    prepend_window_name = True
    use_windowsapps = True
    fullscreen_game_detection = True
    disallowed_chars = ["\\", "/", ":", '*',
                        "?", '"', "<", ">", "|", ".exe", "$"]

    def __init__(self) -> None:
        pass

    def event_handler(self, event, *_) -> None:
        """
        Internal, handles OBS events
        """
        if event == o.OBS_FRONTEND_EVENT_REPLAY_BUFFER_SAVED:
            try:
                self.move_video()
            except BaseException as e:
                print(e)

    def get_last_replay_path(self) -> str:
        """
        Retrieve last replay buffer output
        """

        replay_buffer = o.obs_frontend_get_replay_buffer_output()
        cd = o.calldata_create()
        ph = o.obs_output_get_proc_handler(replay_buffer)

        o.proc_handler_call(
            handler=ph,
            name="get_last_replay",
            params=cd
        )
        path = o.calldata_string(data=cd, name="path")

        o.obs_output_release(replay_buffer)
        return path

    def get_focused_window_name(self) -> str:
        """
        Uses the win32api to grab the name of the currently focused window
        """

        # Get the window text
        w_text = wgui.GetWindowText(wgui.GetForegroundWindow())

        # Sanitize window text
        for char in self.disallowed_chars:
            if char in w_text:
                w_text = w_text.replace(char, "")

        return w_text.strip()

    def is_window_fullscreen(self, hwnd) -> bool:
        """
        Uses the win32api to check if a window is fullscreen
        """

        # Get the window rect
        rect = wgui.GetWindowRect(hwnd)

        # Get the screen size
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)

        # Check if the window is fullscreen
        return rect[2] - rect[0] == screen_width and rect[3] - rect[1] == screen_height

    def get_focused_window_executable_path(self) -> str:
        """
        Uses the win32api and psutil to grab the executable name of the currently focused window
        """

        # Get the handle for the foreground window
        hwnd = wgui.GetForegroundWindow()

        # Get the process ID of the window
        _, pid = win32process.GetWindowThreadProcessId(hwnd)

        # Get the process path from the process ID
        process = psutil.Process(pid)

        return process.exe()

    def get_focused_application_name(self) -> str:
        """
        Uses the win32api to grab the name of the currently focused application
        With help from StackOverflow: https://stackoverflow.com/a/31119785
        """

        exe_path = self.get_focused_window_executable_path()

        try:
            language, codepage = win32api.GetFileVersionInfo(
                exe_path, '\\VarFileInfo\\Translation')[0]
            stringFileInfo = u'\\StringFileInfo\\%04X%04X\\%s' % (
                language, codepage, "FileDescription")
            application_name = win32api.GetFileVersionInfo(
                exe_path, stringFileInfo)
        except:
            application_name = ""

        # Sanitize text
        for char in self.disallowed_chars:
            if char in application_name:
                application_name = application_name.replace(char, "")

        return application_name.strip()

    def move_video(self) -> None:
        """
        Moves the last file the Replay Buffer created\n
        into `base_dir\\{cur_win_name}\\base_name`
        """
        # Get replay path and focused window
        lr_orig_fullpath = self.get_last_replay_path()
        lr_orig_dir, lr_fname = os.path.split(lr_orig_fullpath)

        # First try to get the name from the executable's version information resource
        # If that fails, try to get the name from the window text
        sub_dir = self.get_focused_application_name()
        sub_dir = sub_dir if sub_dir != "" else self.get_focused_window_name()

        if not sub_dir:
            if self.use_windowsapps:
                sub_dir = "Windowsapps"
            elif self.base_dir:
                os.rename(lr_orig_fullpath, os.path.join(self.base_dir, lr_fname))
                return
            return

        if self.fullscreen_game_detection:
            if not self.is_window_fullscreen(wgui.GetForegroundWindow()):
                sub_dir = "Desktop"

        if self.prepend_window_name:
            lr_fname = lr_fname.replace("Replay", sub_dir)

        # Parse out directories
        if self.base_dir:
            lr_base_dir = self.base_dir
        else:
            lr_base_dir = lr_orig_dir

        lr_dir = os.path.join(lr_base_dir, sub_dir)
        if not os.path.exists(lr_dir):
            os.mkdir(lr_dir)

        os.rename(lr_orig_fullpath, os.path.join(lr_dir, lr_fname))


inst = ReplayBufferXtender()


def script_description() -> str:
    return """
    <center><h2>ReplayBufferXtender</h2>
    <p>This is an OBS Python script that automatically renames video files generated 
    by the Replay Buffer based on the window in focus when they are created using the win32api.
    Attempts to emulate the file naming conventions of Nvidia Shadowplay as closely as possible!</p>

    <p><br><b>Author: Myssto</b>
    <br><a href=https://github.com/myssto/OBSReplayBufferXtender/tree/main#obsreplaybufferxtender>Github Documentation</a></p>
    </center>
    """


def script_load(_) -> None:
    o.obs_frontend_add_event_callback(on_event)


def script_unload() -> None:
    o.obs_frontend_remove_event_callback(on_event)


def script_properties() -> any:
    p = o.obs_properties_create()

    # Option to overide existing OBS Recording output path for Replay Buffer only
    o.obs_properties_add_path(
        props=p,
        name="baseSavePath",
        description="Base Save Path",
        type=o.OBS_PATH_DIRECTORY,
        filter=None,
        default_path=None
    )

    # Option to send unknown window names to a Windowsapps directory
    o.obs_properties_add_bool(
        props=p,
        name="useWindowsapps",
        description="Use Windowsapps for Unknown Programs"
    )


    # Option to prepend the window name to the replay file like Shadowplay
    # Will replace the default "Replay" text from the file name
    o.obs_properties_add_bool(
        props=p,
        name="prependWindowName",
        description="Prepend Window Name"
    )

    # Option to detect whether the focused window is fullscreen, and if it isn't, put it in a Desktop directory
    o.obs_properties_add_bool(
        props=p,
        name="fullscreenGameDetection",
        description="Fullscreen Game Detection"
    )

    return p


def script_defaults(s) -> None:
    o.obs_data_set_default_bool(
        data=s,
        name="useWindowsapps",
        val=ReplayBufferXtender.use_windowsapps
    )

    o.obs_data_set_default_bool(
        data=s,
        name="prependWindowName",
        val=ReplayBufferXtender.prepend_window_name
    )

    o.obs_data_set_default_bool(
        data=s,
        name="fullscreenGameDetection",
        val=ReplayBufferXtender.fullscreen_game_detection
    )


def script_update(s) -> None:
    inst.base_dir = o.obs_data_get_string(s, "baseSavePath")
    inst.use_windowsapps = o.obs_data_get_bool(s, "useWindowsapps")
    inst.prepend_window_name = o.obs_data_get_bool(s, "prependWindowName")
    inst.fullscreen_game_detection = o.obs_data_get_bool(s, "fullscreenGameDetection")


def on_event(event, *_) -> None:
    """
    Called by OBS when events are fired\n
    This exists because you cannot bind instance methods as callbacks in OBS
    """

    inst.event_handler(event)
