using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace CellMate
{
    // A thread-only hook in this Excel instance. No global hook, clipboard, or keystroke log.
    internal sealed class EditShortcut : NativeWindow, IDisposable
    {
        private const int DispatchMessage = 0x8000 + 127;
        private const int KeyDown = 0x100, KeyUp = 0x101;
        private readonly HookProc callback;
        private IntPtr hook;
        private int shortcut;
        private int swallowedKey;
        private Request pending;
        private CommitRequest committing;
        private CommitRequest queuedCommit;
        private readonly Timer commitTimer;
        private bool disposed;

        public EditShortcut()
        {
            CreateHandle(new CreateParams { Caption = "CellMate edit shortcut", Parent = new IntPtr(-3) });
            callback = KeyboardHook;
            hook = SetWindowsHookEx(2 /* WH_KEYBOARD */, callback, IntPtr.Zero, GetCurrentThreadId());
            if (hook == IntPtr.Zero) { DestroyHandle(); throw new Win32Exception(Marshal.GetLastWin32Error()); }
            commitTimer = new Timer { Interval = 25 };
            commitTimer.Tick += CheckCommit;
        }

        public void Configure(Preferences preferences)
        {
            shortcut = preferences.Shortcuts.Single(b => b.Command == "Duplicate").Keys;
            CancelCommit();
            if (queuedCommit != null) queuedCommit.Cancelled = true;
        }

        private IntPtr KeyboardHook(int code, IntPtr key, IntPtr flags)
        {
            try
            {
                int vk = unchecked((int)key.ToInt64());
                bool released = (flags.ToInt64() & 0x80000000L) != 0;
                if (code == 0 && !released && vk == 0x1b) CancelCommit();
                if (code == 0 && swallowedKey == vk)
                {
                    if (released) swallowedKey = 0;
                    return new IntPtr(1); // Includes auto-repeat until this physical key is released.
                }
                if (code == 0 && !released && queuedCommit != null && !IsModifier(vk)) queuedCommit.Cancelled = true;
                if (code == 0 && !released && !disposed && pending == null && committing == null && shortcut != 0 &&
                    vk == (shortcut & (int)Keys.KeyCode) && CurrentModifiers() == (shortcut & (int)Keys.Modifiers) &&
                    !Pressed(0x5b) && !Pressed(0x5c))
                {
                    IntPtr editor = GetFocus();
                    if (IsCurrentEditor(editor) && TryReadEdit(out var snapshot))
                    {
                        pending = new Request { Editor = editor, Snapshot = snapshot };
                        if (!PostMessage(Handle, DispatchMessage, IntPtr.Zero, IntPtr.Zero)) { pending = null; return CallNextHookEx(hook, code, key, flags); }
                        swallowedKey = vk;
                        return new IntPtr(1);
                    }
                }
            }
            catch { /* Never let managed exceptions escape a native input callback. */ }
            return CallNextHookEx(hook, code, key, flags);
        }

        protected override void WndProc(ref Message message)
        {
            if (message.Msg != DispatchMessage) { base.WndProc(ref message); return; }
            var request = pending;
            try
            {
                if (disposed || request == null || shortcut == 0 || !IsCurrentEditor(request.Editor) ||
                    !TryReadEdit(out var current) || !current.SameAs(request.Snapshot)) return;
                if (current.SelectionLength > 0) DuplicateText(request.Editor, current);
                else CommitAndDuplicateRows();
            }
            catch (Exception ex) { CancelCommit(); AddIn.ShowError("Could not duplicate the current edit.\n" + ex.Message); }
            finally { pending = null; }
        }

        private static void DuplicateText(IntPtr editor, EditSnapshot snapshot)
        {
            if (snapshot.Text.Length + snapshot.SelectionLength > 32767)
                throw new InvalidOperationException("The duplicated text would exceed Excel's 32,767-character cell limit.");
            WithPlainKeyboard(() =>
            {
                // Let Excel collapse the selection natively, preserving existing rich text.
                SendKey(editor, 0x27 /* Right */);
                if (!TryReadEdit(out var collapsed) || collapsed.Text != snapshot.Text ||
                    collapsed.SelectionLength != 0 || collapsed.Caret != snapshot.SelectionEnd)
                    throw new InvalidOperationException("Excel could not position the caret after the selected text. The text was not duplicated.");
            });
            if (!IsCurrentEditor(editor)) return;
            InsertUnicode(snapshot.SelectedText);
        }

        private static void InsertUnicode(string text)
        {
            // Excel's custom editor does not process synthetic WM_CHAR like a standard Edit control.
            // Use real Unicode input, releasing/restoring only currently held modifiers in one batch.
            var inputs = new System.Collections.Generic.List<Input>();
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch == '\r' || ch == '\n')
                {
                    // Excel ignores a Unicode LF packet; its native in-cell line break is Alt+Enter.
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    inputs.Add(Input.Key(0xa4, false));
                    inputs.Add(Input.Key(0x0d, false)); inputs.Add(Input.Key(0x0d, true));
                    inputs.Add(Input.Key(0xa4, true));
                }
                else { inputs.Add(Input.Unicode(ch, false)); inputs.Add(Input.Unicode(ch, true)); }
            }
            SendPlainInput(inputs);
        }

        private static void SendPlainInput(System.Collections.Generic.IEnumerable<Input> content)
        {
            var inputs = new System.Collections.Generic.List<Input>();
            var held = new[] { 0xa0, 0xa1, 0xa2, 0xa3, 0xa4, 0xa5 }.Where(Pressed).ToArray();
            foreach (int key in held) inputs.Add(Input.Key(key, true));
            inputs.AddRange(content);
            foreach (int key in held) inputs.Add(Input.Key(key, false));
            uint sent = SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(Input)));
            if (sent != inputs.Count)
            {
                int error = Marshal.GetLastWin32Error();
                var restore = held.Select(key => Input.Key(key, false)).ToArray();
                if (restore.Length > 0) SendInput((uint)restore.Length, restore, Marshal.SizeOf(typeof(Input)));
                throw new Win32Exception(error, "Windows could not complete the input. Review the current cell text before continuing.");
            }
        }

        private void CommitAndDuplicateRows()
        {
            // Read the anchor before Enter; Excel may move its selection after committing.
            // If the object model is unavailable, fail before changing or committing the edit.
            dynamic app = AddIn.Application;
            committing = new CommitRequest { Selection = app.Selection, Sheet = app.ActiveSheet, Workbook = app.ActiveWorkbook, Started = Environment.TickCount };
            SendPlainInput(new[] { Input.Key(0x0d, false), Input.Key(0x0d, true) });
            commitTimer.Start();
        }

        private void CheckCommit(object sender, EventArgs args)
        {
            try { CheckCommitCore(); }
            catch { CancelCommit(); } // Excel may be shutting down or rejecting automation.
        }

        private void CheckCommitCore()
        {
            var request = committing;
            if (disposed || request == null) { CancelCommit(); return; }
            var info = new XlCall.FmlaInfo();
            if (XlCall.PenHelper(XlCall.xlGetFmlaInfo, ref info) != 0) { CancelCommit(); return; }
            if (info.wPointMode != XlCall.xlModeReady)
            {
                // Never wait through a validation dialog or leave a future row action armed.
                if (!IsEditor(GetFocus()) || unchecked(Environment.TickCount - request.Started) > 750) CancelCommit();
                return;
            }
            CancelCommit();
            // A validation dialog can report Ready while its modal loop is still active.
            // Only schedule row work after focus returns to this Excel workbook window.
            IntPtr root = GetAncestor(GetFocus(), 2 /* GA_ROOT */);
            if (WindowClass(root) != "XLMAIN" || GetForegroundWindow() != root) return;
            dynamic app = AddIn.Application;
            request.ExpectedSelection = (string)app.Selection.Address;
            queuedCommit = request;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    if (disposed || request.Cancelled) return;
                    if (!Equals((object)app.ActiveWorkbook, request.Workbook) || !Equals((object)app.ActiveSheet, request.Sheet)) return;
                    if ((string)app.Selection.Address != request.ExpectedSelection) return;
                    ((dynamic)request.Selection).Select();
                    Commands.Duplicate();
                }
                catch (Exception ex) { AddIn.ShowError("The edit was committed, but its rows could not be duplicated.\n" + ex.Message); }
                finally { if (queuedCommit == request) queuedCommit = null; }
            });
        }

        private void CancelCommit() { committing = null; commitTimer?.Stop(); }

        private static void WithPlainKeyboard(Action action)
        {
            var saved = new byte[256];
            if (!GetKeyboardState(saved)) throw new Win32Exception(Marshal.GetLastWin32Error());
            var plain = (byte[])saved.Clone();
            foreach (int vk in new[] { 0x10, 0x11, 0x12, 0xa0, 0xa1, 0xa2, 0xa3, 0xa4, 0xa5 }) plain[vk] = 0;
            if (!SetKeyboardState(plain)) throw new Win32Exception(Marshal.GetLastWin32Error());
            try { action(); }
            finally { SetKeyboardState(saved); }
        }

        private static void SendKey(IntPtr editor, int key)
        {
            SendMessage(editor, KeyDown, new IntPtr(key), new IntPtr(1));
            SendMessage(editor, KeyUp, new IntPtr(key), new IntPtr(unchecked((int)0xc0000001)));
        }
        private static int CurrentModifiers() => (Pressed(0x11) ? (int)Keys.Control : 0) | (Pressed(0x10) ? (int)Keys.Shift : 0) | (Pressed(0x12) ? (int)Keys.Alt : 0);
        private static bool Pressed(int key) => (GetKeyState(key) & 0x8000) != 0;
        private static bool IsModifier(int key) => key == 0x10 || key == 0x11 || key == 0x12 || (key >= 0xa0 && key <= 0xa5);
        private static bool IsCurrentEditor(IntPtr editor) => GetFocus() == editor && IsEditor(editor) && GetForegroundWindow() == GetAncestor(editor, 2 /* GA_ROOT */);
        private static bool IsEditor(IntPtr hwnd)
            => WindowClass(hwnd) == "EXCEL6" || WindowClass(hwnd) == "EXCEL<";
        private static string WindowClass(IntPtr hwnd)
        {
            var name = new StringBuilder(64);
            GetClassName(hwnd, name, name.Capacity);
            return name.ToString();
        }
        private static bool TryReadEdit(out EditSnapshot snapshot)
        {
            snapshot = null;
            var info = new XlCall.FmlaInfo();
            if (XlCall.PenHelper(XlCall.xlGetFmlaInfo, ref info) != 0 || info.wPointMode == XlCall.xlModeReady ||
                info.cch < 0 || info.cch > 32767 || (info.cch > 0 && info.lpch == IntPtr.Zero) || info.ichFirst < 0 || info.ichLast < 0) return false;
            snapshot = new EditSnapshot(info.cch == 0 ? "" : Marshal.PtrToStringUni(info.lpch, info.cch), info.ichFirst, info.ichLast, info.ichCaret);
            return true;
        }
        public void Dispose()
        {
            disposed = true;
            shortcut = 0;
            pending = null;
            if (queuedCommit != null) queuedCommit.Cancelled = true;
            CancelCommit();
            commitTimer.Dispose();
            if (hook != IntPtr.Zero) { UnhookWindowsHookEx(hook); hook = IntPtr.Zero; }
            DestroyHandle();
            GC.KeepAlive(callback);
        }
        private sealed class Request { public IntPtr Editor; public EditSnapshot Snapshot; }
        private sealed class CommitRequest { public object Selection, Sheet, Workbook; public string ExpectedSelection; public int Started; public bool Cancelled; }
        private delegate IntPtr HookProc(int code, IntPtr key, IntPtr flags);
        [StructLayout(LayoutKind.Sequential)]
        private struct Input
        {
            public uint Type;
            public InputData Data;
            public static Input Key(int key, bool up) => new Input { Type = 1, Data = new InputData { Keyboard = new KeyboardInput { VirtualKey = (ushort)key, Flags = up ? 2u : 0u } } };
            public static Input Unicode(char ch, bool up) => new Input { Type = 1, Data = new InputData { Keyboard = new KeyboardInput { ScanCode = ch, Flags = 4u | (up ? 2u : 0u) } } };
        }
        [StructLayout(LayoutKind.Explicit)] private struct InputData { [FieldOffset(0)] public KeyboardInput Keyboard; [FieldOffset(0)] public MouseInput Mouse; }
        [StructLayout(LayoutKind.Sequential)] private struct KeyboardInput { public ushort VirtualKey, ScanCode; public uint Flags, Time; public UIntPtr Extra; }
        [StructLayout(LayoutKind.Sequential)] private struct MouseInput { public int X, Y; public uint Data, Flags, Time; public UIntPtr Extra; }
        [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint count, Input[] inputs, int size);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetWindowsHookEx(int type, HookProc callback, IntPtr module, uint thread);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr key, IntPtr flags);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern IntPtr GetFocus();
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern short GetKeyState(int key);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool GetKeyboardState(byte[] keys);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool SetKeyboardState(byte[] keys);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int maxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool PostMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
    }

    internal sealed class EditSnapshot
    {
        public readonly string Text;
        public readonly int SelectionStart, SelectionEnd, Caret;
        public int SelectionLength => SelectionEnd - SelectionStart;
        public string SelectedText => Text.Substring(SelectionStart, SelectionLength);
        public EditSnapshot(string text, int start, int end, int caret)
        {
            Text = text;
            SelectionStart = Math.Min(text.Length, Math.Min(start, end));
            SelectionEnd = Math.Min(text.Length, Math.Max(start, end));
            Caret = Math.Max(0, Math.Min(text.Length, caret));
        }
        public bool SameAs(EditSnapshot other) => other != null && Text == other.Text && SelectionStart == other.SelectionStart && SelectionEnd == other.SelectionEnd && Caret == other.Caret;
    }
}
