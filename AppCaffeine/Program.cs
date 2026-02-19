using System.Runtime.InteropServices;
using System.Diagnostics;
using ObjCRuntime;
using ServiceManagement;

// Initialize the native macOS Application
NSApplication.Init();
var app = NSApplication.SharedApplication;
var delegateInstance = new AppDelegate();
app.Delegate = delegateInstance;
app.Run();

public class AppDelegate : NSApplicationDelegate
{
    private NSStatusItem? _statusItem;
    private CancellationTokenSource? _cts;
    private bool _running;

    // Current Selections
    private int? _targetPid;
    private string? _targetWindowName;
    private string _targetAppName = "None";

    public override void DidFinishLaunching(NSNotification notification)
    { 
        LoadSettings(); // <--- Add this
        _statusItem = NSStatusBar.SystemStatusBar.CreateStatusItem(NSStatusItemLength.Variable);
        _statusItem.Button.Title = "⌨️";
        RefreshMenu();
    }
    private NSImage ResizeImage(NSImage sourceImage, CGSize newSize)
    {
        var newImage = new NSImage(newSize);
        newImage.LockFocus();
        sourceImage.Size = newSize;
        if (NSGraphicsContext.CurrentContext != null)
            NSGraphicsContext.CurrentContext.ImageInterpolation = NSImageInterpolation.High;
        sourceImage.Draw(new CGPoint(0, 0), new CGRect(0, 0, newSize.Width, newSize.Height), NSCompositingOperation.SourceOver, 1.0f);
        newImage.UnlockFocus();
        return newImage;
    }
    private void RefreshMenu()
    {
        var menu = new NSMenu();

        // 1. Status & Toggle
        var statusHeader = new NSMenuItem($"Target: {_targetAppName} > {(_targetWindowName ?? "Any")}");
        statusHeader.Enabled = false;
        menu.AddItem(statusHeader);

        var toggleItem = new NSMenuItem(_running ? "Stop Automation" : "Start Automation");
        toggleItem.Activated += (_, _) => {
            _running = !_running;
            if (_running) StartLoop(); else StopLoop();
            RefreshMenu();
            SaveSettings();
        };
        menu.AddItem(toggleItem);
        menu.AddItem(NSMenuItem.SeparatorItem);

        // 2. App Selection Submenu
        var appsMenu = new NSMenuItem("Select App");
        var subApps = new NSMenu();
        
        // Get apps with windows and a GUI
        var apps = NSWorkspace.SharedWorkspace.RunningApplications
            .Where(a => a.ActivationPolicy == NSApplicationActivationPolicy.Regular);

        foreach (var app in apps)
        {
            var appItem = new NSMenuItem(app.LocalizedName!, (_, _) => {
                _targetPid = app.ProcessIdentifier;
                _targetAppName = app.LocalizedName!;
                _targetWindowName = null;
                SaveSettings();
                RefreshMenu();
            });

            // Add and resize the icon
            if (app.Icon != null)
            {
                appItem.Image = ResizeImage(app.Icon, new CGSize(18, 18));
            }
            // Inside the App loop:
            if (_targetPid == app.ProcessIdentifier) 
                appItem.State = NSCellStateValue.On; // Shows a checkmark

            subApps.AddItem(appItem);
        }
        appsMenu.Submenu = subApps;
        menu.AddItem(appsMenu);

        // 3. Window Selection Submenu (Only if app selected)
        if (_targetPid.HasValue)
        {
            var winMenu = new NSMenuItem("Select Window");
            var subWin = new NSMenu();
            var windowTitles = GetWindowTitles(_targetPid.Value);

            foreach (var title in windowTitles)
            {
                var winItem = new NSMenuItem(title, (_, _) =>
                {
                    _targetWindowName = title;
                    SaveSettings();
                    RefreshMenu();
                });

                // Inside the Window loop:
                if (_targetWindowName == title) 
                    winItem.State = NSCellStateValue.On; // Shows a checkmark
                subWin.AddItem(winItem);
            }
            winMenu.Submenu = subWin;
            
            menu.AddItem(winMenu);
        }

        menu.AddItem(NSMenuItem.SeparatorItem);
        menu.AddItem(new NSMenuItem("Quit", (_, _) => NSApplication.SharedApplication.Terminate(this)));
        
        _statusItem!.Menu = menu;
        
        var launchItem = new NSMenuItem("Launch at Login", (s, _) => {
            ToggleLaunchAtLogin((NSMenuItem)s!);
        });

        // Set initial checkmark state based on system status
        launchItem.State = IsLaunchAtLoginEnabled() ? NSCellStateValue.On : NSCellStateValue.Off;

        menu.AddItem(launchItem);
        menu.AddItem(NSMenuItem.SeparatorItem);
    }
    
    private void SaveSettings()
    {
        var defaults = NSUserDefaults.StandardUserDefaults;
    
        // Find the bundle ID for the target PID
        var app = NSRunningApplication.GetRunningApplication(_targetPid ?? 0);
        if (app.BundleIdentifier != null)
        {
            defaults.SetString(app.BundleIdentifier, "TargetBundleId");
            defaults.SetString(_targetAppName, "TargetAppName");
        }
    
        if (!string.IsNullOrEmpty(_targetWindowName))
            defaults.SetString(_targetWindowName, "TargetWindowName");
        else
            defaults.RemoveObject("TargetWindowName");

        defaults.SetBool(_running, "Running");
        defaults.Synchronize();
    }

    private void LoadSettings()
    {
        var defaults = NSUserDefaults.StandardUserDefaults;
        var bundleId = defaults.StringForKey("TargetBundleId");
        var winName = defaults.StringForKey("TargetWindowName");
        var appName = defaults.StringForKey("TargetAppName");
        _running = defaults.BoolForKey("Running");
        if (_running)
        {
            StartLoop();
        }

        if (string.IsNullOrEmpty(bundleId)) return;
        // Try to find if that app is currently running to get its PID
        var runningApp = NSRunningApplication.GetRunningApplications(bundleId).FirstOrDefault();
        if (runningApp != null)
        {
            _targetPid = runningApp.ProcessIdentifier;
            _targetAppName = runningApp.LocalizedName!;
        }
        else
        {
            _targetAppName = appName ?? "None (Not Running)";
        }
        _targetWindowName = winName;
    }
    private bool IsLaunchAtLoginEnabled()
    {
        // Check if the service is already registered
        return SMAppService.MainApp.Status == SMAppServiceStatus.Enabled;
    }

    private void ToggleLaunchAtLogin(NSMenuItem item)
    {
        var service = SMAppService.MainApp;
        NSError? error;

        if (service.Status == SMAppServiceStatus.Enabled)
        {
            service.Unregister(out error);
            item.State = NSCellStateValue.Off;
        }
        else
        {
            service.Register(out error);
            item.State = NSCellStateValue.On;
        }

        if (error != null)
        {
            Console.WriteLine($"Login Item Error: {error.LocalizedDescription}");
        }
    }
    private List<string> GetWindowTitles(int pid)
    {
        var titles = new List<string>();
        IntPtr appRef = WindowAutomator.AXUIElementCreateApplication(pid);

        if (WindowAutomator.AXUIElementCopyAttributeValue(appRef, ((NSString)"AXWindows").Handle, out var windowListPtr) == 0)
        {
            var windows = Runtime.GetNSObject<NSArray>(windowListPtr);
            if (windows == null) return titles;
            for (nuint i = 0; i < windows.Count; i++)
            {
                IntPtr winRef = windows.ValueAt(i);
                WindowAutomator.AXUIElementCopyAttributeValue(winRef, ((NSString)"AXTitle").Handle, out var titlePtr);
                var t = ((NSString)Runtime.GetNSObject(titlePtr)!).ToString();
                if (!string.IsNullOrEmpty(t)) titles.Add(t);
            }
        }
        return titles;
    }

    private void StartLoop()
    {
        if (!_targetPid.HasValue) return;
        _cts = new CancellationTokenSource();

        Task.Run(async () => {
            while (!_cts.Token.IsCancellationRequested)
            {
                // If the process died or restarted, try to find it again by name
                var process = Process.GetProcessesByName(_targetAppName).FirstOrDefault();
                if (process != null && _targetWindowName != null)
                {
                    WindowAutomator.SendKeyToSpecificWindow(process.Id, _targetWindowName, 0x5A); //F20 key
                }
                await Task.Delay(TimeSpan.FromMinutes(5));
            }
        });
    }

    private void StopLoop() => _cts?.Cancel();
}

public static class WindowAutomator
{
    private const string HiServices = "/System/Library/Frameworks/ApplicationServices.framework/Frameworks/HIServices.framework/HIServices";
    [DllImport("/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation")]
    private static extern void CFRelease(IntPtr obj);

    [DllImport("/System/Library/Frameworks/ApplicationServices.framework/Frameworks/HIServices.framework/HIServices")]
    public static extern IntPtr AXUIElementCreateApplication(int pid);

    [DllImport("/System/Library/Frameworks/ApplicationServices.framework/Frameworks/HIServices.framework/HIServices")]
    public static extern int AXUIElementCopyAttributeValue(IntPtr element, IntPtr attribute, out IntPtr value);

    [DllImport("/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics")]
    private static extern IntPtr CGEventCreateKeyboardEvent(IntPtr source, ushort virtualKey, bool keyDown);

    [DllImport("/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics")]
    private static extern void CGEventPostToPid(int pid, IntPtr eventRef);

    [DllImport(HiServices)]
    private static extern int AXUIElementPerformAction(IntPtr element, IntPtr action);


    public static void SendKeyToSpecificWindow(int targetPid, string? windowTitle, ushort keyCode)
    {
        // 1. Capture the currently active app BEFORE we do anything
        var previousApp = NSWorkspace.SharedWorkspace.FrontmostApplication;

        IntPtr appRef = AXUIElementCreateApplication(targetPid);
        IntPtr windowListPtr;
    
        if (AXUIElementCopyAttributeValue(appRef, ((NSString)"AXWindows").Handle, out windowListPtr) == 0)
        {
            var windows = Runtime.GetNSObject<NSArray>(windowListPtr);
            if (windows != null)
                for (nuint i = 0; i < windows.Count; i++)
                {
                    IntPtr winRef = windows.ValueAt(i);
                    AXUIElementCopyAttributeValue(winRef, ((NSString)"AXTitle").Handle, out var titlePtr);
                    var currentTitle = ((NSString)Runtime.GetNSObject(titlePtr)!).ToString();

                    if (string.IsNullOrEmpty(windowTitle) || currentTitle.Contains(windowTitle))
                    {
                        // 2. Focus the target window
                        AXUIElementPerformAction(winRef, ((NSString)"AXRaise").Handle);

                        // 3. Send the key via CGEvent
                        IntPtr keyDown = CGEventCreateKeyboardEvent(IntPtr.Zero, keyCode, true);
                        CGEventPostToPid(targetPid, keyDown);
                        CFRelease(keyDown);

                        IntPtr keyUp = CGEventCreateKeyboardEvent(IntPtr.Zero, keyCode, false);
                        CGEventPostToPid(targetPid, keyUp);
                        CFRelease(keyUp);

                        // 4. RESTORE FOCUS: Bring the previous app back to the front
                        // NSApplicationActivateOptions.AllWindows ensures all windows of the previous app return
                        previousApp.Activate(NSApplicationActivationOptions.ActivateAllWindows);

                        break;
                    }
                }
        }
    }
}
