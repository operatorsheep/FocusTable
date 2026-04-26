using System;
using System.Collections.Generic;
using UnityEngine;
using KSP.UI.Screens;

namespace FocusTable
{
    [KSPAddon(KSPAddon.Startup.FlightAndKSC, false)]
    public class FocusTableMod : MonoBehaviour
    {
        // Window state
        private bool    _showWindow = false;
        private Rect    _windowRect = new Rect(Screen.width - 360, 40, 340, 560);
        private Vector2 _scrollPos  = Vector2.zero;
        private int     _windowId;

        // Resize
        private bool    _isResizing      = false;
        private Vector2 _resizeMouseStart;
        private Vector2 _resizeOrigSize;
        private const float MIN_WIN_WIDTH  = 260f;
        private const float MIN_WIN_HEIGHT = 160f;

        // Toolbar button
        private ApplicationLauncherButton _appButton;
        private Texture2D _buttonTexture;

        // Data
        private List<VesselEntry> _entries = new List<VesselEntry>();
        private float _lastRefresh = 0f;
        private const float REFRESH_INTERVAL = 1.0f;
        private RendezvousInfo _rdv;

        // Expanded vessels (showing docking ports)
        private HashSet<Vessel> _expandedVessels = new HashSet<Vessel>();

        // Filters
        private string _filterText = "";
        private HashSet<string> _activeCategories = new HashSet<string>(); // left-click: show only
        private HashSet<string> _hiddenCategories  = new HashSet<string>(); // right-click: hide

        // Sorting
        private enum SortColumn { Distance, Name, Type }
        private SortColumn _sortColumn    = SortColumn.Distance;
        private bool       _sortAscending = true;

        // Favourites (vessel GUIDs)
        private HashSet<string> _favouriteIds = new HashSet<string>();

        // Notes keyed by vessel GUID
        private Dictionary<string, string> _vesselNotes = new Dictionary<string, string>();

        // Deferred action — executed after DrawWindow returns, avoids modifying _entries mid-foreach
        private Action _deferredAction = null;

        // Context menu
        private bool        _contextMenuOpen = false;
        private VesselEntry _contextMenuEntry;
        private Rect        _contextMenuRect;
        private int         _contextMenuWindowId;

        // Notes popup
        private bool        _notesOpen     = false;
        private VesselEntry _notesEntry;
        private string      _notesId       = "";
        private string      _notesEditText = "";
        private Rect        _notesRect;
        private int         _notesWindowId;

        // Rename popup
        private bool        _renameOpen  = false;
        private VesselEntry _renameEntry;
        private string      _renameText  = "";
        private Rect        _renameRect;
        private int         _renameWindowId;

        // Terminate confirmation
        private bool        _terminateOpen  = false;
        private VesselEntry _terminateEntry;
        private Rect        _terminateRect;
        private int         _terminateWindowId;

        // Category definitions: display label, type string to match
        private static readonly (string label, string type)[] Categories = new[]
        {
            ("Shp", "Ship"),
            ("Prb", "Probe"),
            ("Sta", "Station"),
            ("Bas", "Base"),
            ("Lnd", "Lander"),
            ("Rov", "Rover"),
            ("Dbr", "Debris"),
            ("Pln", "Plane"),
            ("Rly", "Relay"),
            ("EVA", "EVA"),
            ("Bdy", "Body")
        };

        // Styles
        private GUIStyle _headerStyle;
        private GUIStyle _subHeaderStyle;
        private GUIStyle _rowStyleA;
        private GUIStyle _rowStyleB;
        private GUIStyle _activeStyle;
        private GUIStyle _targetedStyle;
        private GUIStyle _portRowStyle;
        private GUIStyle _portOccupiedStyle;
        private GUIStyle _filterStyle;
        private GUIStyle _categoryBtnOn;
        private GUIStyle _categoryBtnOff;
        private GUIStyle _categoryBtnHide;
        private GUIStyle _smallLabelStyle;
        private GUIStyle _colHeaderStyle;
        private GUIStyle _resizeGripStyle;
        private GUIStyle _popupBtnStyle;
        private GUIStyle _dangerBtnStyle;
        private bool _stylesInitialized = false;

        private struct VesselEntry
        {
            public string Name;
            public string Type;
            public string Situation;
            public double Distance;
            public Vessel Vessel;
            public CelestialBody Body;
            public bool IsActive;
            public bool IsTargeted;
        }

        private struct PortInfo
        {
            public ModuleDockingNode Node;
            public string PartName;
            public string State;        // Ready / Docked / PreAttached
            public bool Occupied;
            public bool Compatible;     // compatible with active vessel's ports
            public float Angle;         // degrees off axis
            public float Roll;          // roll offset
            public float Distance;      // metres
            public float ApproachSpeed; // m/s (closing speed)
        }

        private struct RendezvousInfo
        {
            public bool Valid;
            public string TargetName;
            public double HohmannDv;     // total m/s (dv1 + dv2)
            public double TransferTime;  // seconds for half-orbit transfer
            public double TimeToWindow;  // seconds to next launch window
            public double CurPhase;      // current phase angle, degrees
            public double ReqPhase;      // required phase angle at departure, degrees
            public double RelInclination;// degrees between orbital planes
            public double PlaneChangeDv; // m/s for plane-change burn
        }

        // ── Lifecycle ──────────────────────────────────────────────────────────

        private void Start()
        {
            _windowId            = GUIUtility.GetControlID(GetHashCode(), FocusType.Passive);
            _contextMenuWindowId = _windowId + 1;
            _notesWindowId       = _windowId + 2;
            _renameWindowId      = _windowId + 3;
            _terminateWindowId   = _windowId + 4;
            GameEvents.onGUIApplicationLauncherReady.Add(AddAppButton);
            GameEvents.onGUIApplicationLauncherDestroyed.Add(RemoveAppButton);
            GameEvents.onVesselChange.Add(OnVesselChange);
            LoadData();
        }

        private void OnDestroy()
        {
            GameEvents.onGUIApplicationLauncherReady.Remove(AddAppButton);
            GameEvents.onGUIApplicationLauncherDestroyed.Remove(RemoveAppButton);
            GameEvents.onVesselChange.Remove(OnVesselChange);
            RemoveAppButton();
            InputLockManager.RemoveControlLock("FocusTable");
        }

        private void OnVesselChange(Vessel v)
        {
            // Re-sync button toggle state after KSP switches active vessel
            if (_appButton != null)
            {
                if (_showWindow) _appButton.SetTrue(makeCall: false);
                else             _appButton.SetFalse(makeCall: false);
            }
            RefreshEntries();
        }

        // ── App Launcher ───────────────────────────────────────────────────────

        private void AddAppButton()
        {
            if (_appButton != null) return;
            _buttonTexture = CreateButtonTexture();
            _appButton = ApplicationLauncher.Instance.AddModApplication(
                onTrue:  () => { _showWindow = true;  RefreshEntries(); },
                onFalse: () => _showWindow = false,
                onHover: null, onHoverOut: null,
                onEnable: null, onDisable: null,
                visibleInScenes: ApplicationLauncher.AppScenes.FLIGHT | ApplicationLauncher.AppScenes.MAPVIEW,
                texture: _buttonTexture
            );
            // If the launcher was rebuilt mid-session (e.g. vessel switch), restore the toggle state
            if (_showWindow)
                _appButton.SetTrue(makeCall: false);
        }

        private void RemoveAppButton()
        {
            if (_appButton == null) return;
            ApplicationLauncher.Instance?.RemoveModApplication(_appButton);
            _appButton = null;
        }

        private Texture2D CreateButtonTexture()
        {
            var tex = new Texture2D(38, 38, TextureFormat.ARGB32, false);
            var pixels = new Color[38 * 38];
            for (int i = 0; i < pixels.Length; i++) pixels[i] = Color.clear;
            Color line   = new Color(0.9f, 0.9f, 0.9f, 1f);
            Color accent = new Color(0.4f, 0.8f, 1f,  1f);
            DrawRect(pixels, 38, 4, 6,  30, 3, line);
            DrawRect(pixels, 38, 4, 16, 30, 3, accent);
            DrawRect(pixels, 38, 4, 26, 30, 3, line);
            tex.SetPixels(pixels);
            tex.Apply();
            return tex;
        }

        private void DrawRect(Color[] buf, int w, int x, int y, int rw, int rh, Color c)
        {
            for (int dy = 0; dy < rh; dy++)
                for (int dx = 0; dx < rw; dx++)
                {
                    int idx = (y + dy) * w + (x + dx);
                    if (idx >= 0 && idx < buf.Length) buf[idx] = c;
                }
        }

        // ── Update ─────────────────────────────────────────────────────────────

        private void Update()
        {
            if (Input.GetKey(KeyCode.LeftAlt) && Input.GetKeyDown(KeyCode.F))
            {
                _showWindow = !_showWindow;
                if (_appButton != null)
                {
                    if (_showWindow) _appButton.SetTrue(makeCall: false);
                    else             _appButton.SetFalse(makeCall: false);
                }
                if (_showWindow) RefreshEntries();
            }

            if (Input.GetKeyDown(KeyCode.Escape))
            {
                _contextMenuOpen = false;
                _notesOpen       = false;
                _renameOpen      = false;
                _terminateOpen   = false;
            }

            if (_showWindow && Time.realtimeSinceStartup - _lastRefresh > REFRESH_INTERVAL)
            {
                RefreshEntries();
                _lastRefresh = Time.realtimeSinceStartup;
            }

            // Lock KSP input when mouse is over any of our windows to prevent click-through
            Vector2 guiMouse = new Vector2(Input.mousePosition.x, Screen.height - Input.mousePosition.y);
            bool overUI = (_showWindow      && _windowRect.Contains(guiMouse))         ||
                          (_contextMenuOpen && _contextMenuRect.Contains(guiMouse))    ||
                          (_notesOpen       && _notesRect.Contains(guiMouse))          ||
                          (_renameOpen      && _renameRect.Contains(guiMouse))         ||
                          (_terminateOpen   && _terminateRect.Contains(guiMouse));
            if (overUI)
            {
                InputLockManager.SetControlLock(ControlTypes.ALLBUTCAMERAS, "FocusTable");
                Input.ResetInputAxes();
            }
            else
            {
                InputLockManager.RemoveControlLock("FocusTable");
            }
        }

        // ── Data ───────────────────────────────────────────────────────────────

        private void RefreshEntries()
        {
            _entries.Clear();

            var activeVessel  = FlightGlobals.ActiveVessel;
            var currentTarget = activeVessel?.targetObject;

            if (FlightGlobals.Vessels != null)
            {
                foreach (Vessel v in FlightGlobals.Vessels)
                {
                    if (v == null) continue;
                    double dist = activeVessel != null
                        ? Vector3d.Distance(v.GetWorldPos3D(), activeVessel.GetWorldPos3D()) : 0;
                    bool isTargeted = currentTarget is Vessel tv && tv == v;
                    _entries.Add(new VesselEntry
                    {
                        Name = string.IsNullOrEmpty(v.vesselName) ? "(unnamed)" : v.vesselName,
                        Type = v.vesselType.ToString(),
                        Situation = v.situation.ToString().Replace("_", " "),
                        Distance = dist, Vessel = v, Body = null,
                        IsActive = v == activeVessel, IsTargeted = isTargeted
                    });
                }
            }

            if (FlightGlobals.Bodies != null)
            {
                foreach (CelestialBody body in FlightGlobals.Bodies)
                {
                    if (body == null) continue;
                    double dist = activeVessel != null
                        ? Vector3d.Distance(body.position, activeVessel.GetWorldPos3D()) : 0;
                    bool isTargeted = currentTarget is CelestialBody tb && tb == body;
                    _entries.Add(new VesselEntry
                    {
                        Name = body.displayName.LocalizeRemoveGender(),
                        Type = "Body", Situation = "", Distance = dist,
                        Vessel = null, Body = body,
                        IsActive = false, IsTargeted = isTargeted
                    });
                }
            }

            ApplySort();
            _rdv = CalcRendezvous();
        }

        private void ApplySort()
        {
            _entries.Sort((a, b) =>
            {
                if (a.IsActive   && !b.IsActive)   return -1;
                if (!a.IsActive  && b.IsActive)    return 1;
                bool afav = a.Vessel != null && _favouriteIds.Contains(a.Vessel.id.ToString());
                bool bfav = b.Vessel != null && _favouriteIds.Contains(b.Vessel.id.ToString());
                if (afav && !bfav) return -1;
                if (!afav && bfav) return 1;
                if (a.IsTargeted && !b.IsTargeted) return -1;
                if (!a.IsTargeted && b.IsTargeted) return 1;
                int cmp;
                switch (_sortColumn)
                {
                    case SortColumn.Name: cmp = string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase); break;
                    case SortColumn.Type: cmp = string.Compare(a.Type, b.Type, StringComparison.OrdinalIgnoreCase); break;
                    default:              cmp = a.Distance.CompareTo(b.Distance); break;
                }
                return _sortAscending ? cmp : -cmp;
            });
        }

        private string ColLabel(SortColumn col, string label)
            => _sortColumn == col ? label + (_sortAscending ? " ▲" : " ▼") : label;

        private void ClickSort(SortColumn col)
        {
            if (_sortColumn == col) _sortAscending = !_sortAscending;
            else { _sortColumn = col; _sortAscending = true; }
            ApplySort();
        }

        // ── Docking Port Helpers ───────────────────────────────────────────────

        private List<PortInfo> GetPorts(Vessel target)
        {
            var list = new List<PortInfo>();
            if (target == null) return list;

            var activeVessel = FlightGlobals.ActiveVessel;

            var activeSizes = new HashSet<string>();
            if (activeVessel != null && activeVessel.parts != null)
            {
                foreach (var p in activeVessel.parts)
                    foreach (ModuleDockingNode m in p.Modules.GetModules<ModuleDockingNode>())
                        if (!string.IsNullOrEmpty(m.nodeType))
                            activeSizes.Add(m.nodeType);
            }

            if (target.loaded && target.parts != null)
            {
                foreach (var part in target.parts)
                {
                    foreach (ModuleDockingNode node in part.Modules.GetModules<ModuleDockingNode>())
                    {
                        bool occupied = node.state != null &&
                                        (node.state.StartsWith("Docked") || node.state == "PreAttached");
                        bool compatible = activeSizes.Count == 0 ||
                                          string.IsNullOrEmpty(node.nodeType) ||
                                          activeSizes.Contains(node.nodeType);

                        float angle = 0f, roll = 0f, dist = 0f, speed = 0f;
                        if (activeVessel != null)
                        {
                            Vector3 toPort = part.transform.position - activeVessel.transform.position;
                            dist  = toPort.magnitude;
                            speed = (float)Vector3d.Dot(
                                        (activeVessel.GetWorldPos3D() - target.GetWorldPos3D()).normalized,
                                        activeVessel.GetObtVelocity() - target.GetObtVelocity());
                            Vector3 fwd = activeVessel.transform.forward;
                            angle = Vector3.Angle(fwd, toPort.normalized);
                            roll = Vector3.SignedAngle(activeVessel.transform.up, part.transform.up, toPort.normalized);
                        }

                        list.Add(new PortInfo
                        {
                            Node = node,
                            PartName = part.partInfo?.title ?? part.partName,
                            State = node.state ?? "Unknown",
                            Occupied = occupied, Compatible = compatible,
                            Angle = angle, Roll = roll, Distance = dist, ApproachSpeed = speed
                        });
                    }
                }
            }
            else if (target.protoVessel?.protoPartSnapshots != null)
            {
                // Vessel is outside physics range — read from persistent save data
                foreach (var pps in target.protoVessel.protoPartSnapshots)
                {
                    foreach (var ppms in pps.modules)
                    {
                        if (ppms.moduleName != "ModuleDockingNode") continue;
                        string state    = ppms.moduleValues.GetValue("state") ?? "Unknown";
                        string nodeType = ppms.moduleValues.GetValue("nodeType") ?? "";
                        bool occupied   = state.StartsWith("Docked") || state == "PreAttached";
                        bool compatible = activeSizes.Count == 0 ||
                                          string.IsNullOrEmpty(nodeType) ||
                                          activeSizes.Contains(nodeType);
                        list.Add(new PortInfo
                        {
                            Node = null,
                            PartName = pps.partInfo?.title ?? pps.partName,
                            State = state,
                            Occupied = occupied, Compatible = compatible,
                            Angle = 0f, Roll = 0f, Distance = 0f, ApproachSpeed = 0f
                        });
                    }
                }
            }

            return list;
        }

        // ── Rendezvous Calculator ──────────────────────────────────────────────

        private RendezvousInfo CalcRendezvous()
        {
            var av = FlightGlobals.ActiveVessel;
            if (av?.orbit == null) return default;

            var tv = av.targetObject as Vessel;
            if (tv?.orbit == null) return default;

            Orbit o1 = av.orbit;
            Orbit o2 = tv.orbit;
            if (o1.referenceBody != o2.referenceBody) return default;
            if (o1.eccentricity >= 1.0 || o2.eccentricity >= 1.0) return default;

            double mu = o1.referenceBody.gravParameter;
            double r1 = o1.semiMajorAxis;
            double r2 = o2.semiMajorAxis;
            if (r1 <= 0 || r2 <= 0 || o1.period <= 0 || o2.period <= 0) return default;

            // Hohmann transfer dv
            double at  = (r1 + r2) / 2.0;
            double dv1 = Math.Abs(Math.Sqrt(mu * (2.0 / r1 - 1.0 / at)) - Math.Sqrt(mu / r1));
            double dv2 = Math.Abs(Math.Sqrt(mu / r2) - Math.Sqrt(mu * (2.0 / r2 - 1.0 / at)));
            double transferTime = Math.PI * Math.Sqrt(at * at * at / mu);

            // Mean motions
            double n1 = 2.0 * Math.PI / o1.period;
            double n2 = 2.0 * Math.PI / o2.period;

            // Required phase angle: target must be this far ahead at departure
            // so it arrives at the intercept point exactly when we do
            double reqPhase = Math.PI - n2 * transferTime;
            reqPhase = ((reqPhase % (2.0 * Math.PI)) + 2.0 * Math.PI) % (2.0 * Math.PI);

            // Current phase angle (signed, in active vessel's orbital plane)
            Vector3d p1     = av.GetWorldPos3D() - o1.referenceBody.position;
            Vector3d p2     = tv.GetWorldPos3D() - o1.referenceBody.position;
            Vector3d normal = o1.GetOrbitNormal().normalized;
            double   curPhase = Math.Atan2(
                Vector3d.Dot(Vector3d.Cross(p1.normalized, p2.normalized), normal),
                Vector3d.Dot(p1.normalized, p2.normalized));
            if (curPhase < 0) curPhase += 2.0 * Math.PI;

            // Time to next launch window
            double phaseRate = n2 - n1;
            double timeToWindow;
            if (Math.Abs(phaseRate) < 1e-12)
            {
                timeToWindow = double.NaN;
            }
            else
            {
                double synodicPeriod = Math.Abs(2.0 * Math.PI / phaseRate);
                double phaseDiff     = reqPhase - curPhase;
                timeToWindow = ((phaseDiff / phaseRate) % synodicPeriod + synodicPeriod) % synodicPeriod;
            }

            // Relative inclination and plane-change dv
            Vector3d n1v   = o1.GetOrbitNormal().normalized;
            Vector3d n2v   = o2.GetOrbitNormal().normalized;
            double relInc  = Math.Acos(Math.Max(-1.0, Math.Min(1.0, Vector3d.Dot(n1v, n2v)))) * (180.0 / Math.PI);
            double planeDv = 2.0 * Math.Sqrt(mu / r1) * Math.Sin(relInc * Math.PI / 360.0);

            return new RendezvousInfo
            {
                Valid          = true,
                TargetName     = tv.vesselName,
                HohmannDv      = dv1 + dv2,
                TransferTime   = transferTime,
                TimeToWindow   = timeToWindow,
                CurPhase       = curPhase * (180.0 / Math.PI),
                ReqPhase       = reqPhase * (180.0 / Math.PI),
                RelInclination = relInc,
                PlaneChangeDv  = planeDv
            };
        }

        private static string FormatDuration(double seconds)
        {
            if (double.IsNaN(seconds) || double.IsInfinity(seconds) || seconds < 0) return "N/A";
            long s = (long)seconds;
            long h = s / 3600; s %= 3600;
            long m = s / 60;   s %= 60;
            return h > 0 ? $"{h}h{m:D2}m{s:D2}s" : $"{m}m{s:D2}s";
        }

        // ── GUI ────────────────────────────────────────────────────────────────

        private void OnGUI()
        {
            if (!_showWindow) return;
            InitStyles();

            _windowRect = GUILayout.Window(
                _windowId, _windowRect, DrawWindow,
                "FocusTable",
                HighLogic.Skin.window,
                GUILayout.Width(_windowRect.width),
                GUILayout.Height(_windowRect.height)
            );

            _windowRect.x = Mathf.Clamp(_windowRect.x, 0, Screen.width  - _windowRect.width);
            _windowRect.y = Mathf.Clamp(_windowRect.y, 0, Screen.height - _windowRect.height);

            HandleResize();

            if (_deferredAction != null)
            {
                _deferredAction();
                _deferredAction = null;
            }

            // Dismiss context menu on left-click outside it
            if (_contextMenuOpen && Event.current.type == EventType.MouseDown && Event.current.button == 0
                && !_contextMenuRect.Contains(Event.current.mousePosition))
            {
                _contextMenuOpen = false;
            }

            if (_contextMenuOpen)
            {
                _contextMenuRect = GUI.Window(_contextMenuWindowId, _contextMenuRect,
                    DrawContextMenu, "", HighLogic.Skin.window);
                GUI.BringWindowToFront(_contextMenuWindowId);
            }
            if (_notesOpen)
                _notesRect = GUI.Window(_notesWindowId, _notesRect,
                    DrawNotesWindow, "Notes", HighLogic.Skin.window);
            if (_renameOpen)
                _renameRect = GUI.Window(_renameWindowId, _renameRect,
                    DrawRenameWindow, "Rename Vessel", HighLogic.Skin.window);
            if (_terminateOpen)
                _terminateRect = GUI.Window(_terminateWindowId, _terminateRect,
                    DrawTerminateWindow, "Terminate Vessel", HighLogic.Skin.window);
        }

        private void HandleResize()
        {
            var e = Event.current;
            if (e == null) return;

            Rect grip = new Rect(_windowRect.xMax - 24, _windowRect.yMax - 24, 24, 24);

            if (e.type == EventType.MouseDown && e.button == 0 && grip.Contains(e.mousePosition))
            {
                _isResizing      = true;
                _resizeMouseStart = e.mousePosition;
                _resizeOrigSize   = new Vector2(_windowRect.width, _windowRect.height);
                e.Use();
            }

            if (_isResizing)
            {
                if (e.type == EventType.MouseDrag)
                {
                    Vector2 delta   = e.mousePosition - _resizeMouseStart;
                    _windowRect.width  = Mathf.Max(MIN_WIN_WIDTH,  _resizeOrigSize.x + delta.x);
                    _windowRect.height = Mathf.Max(MIN_WIN_HEIGHT, _resizeOrigSize.y + delta.y);
                    e.Use();
                }
                if (e.type == EventType.MouseUp)
                {
                    _isResizing = false;
                    e.Use();
                }
            }
        }

        private void DrawWindow(int id)
        {
            GUILayout.BeginVertical();

            // Stats
            int vesselCount = 0;
            foreach (var e in _entries) if (e.Vessel != null) vesselCount++;
            GUILayout.Label($"<b>{vesselCount} vessels · {_entries.Count - vesselCount} bodies</b>  (Alt+F)", _headerStyle);

            // Text filter
            GUILayout.BeginHorizontal();
            GUILayout.Label("Search:", GUILayout.Width(48));
            _filterText = GUILayout.TextField(_filterText, _filterStyle);
            if (GUILayout.Button("X", GUILayout.Width(22)))
                _filterText = "";
            GUILayout.EndHorizontal();

            GUILayout.Space(2);

            // Category emoji toggle buttons — compact, 2 rows of ~6
            GUILayout.BeginHorizontal();
            GUILayout.Label("<b>Filter:</b>", _headerStyle, GUILayout.Width(42));
            if (GUILayout.Button("X all", _categoryBtnOff, GUILayout.Width(36), GUILayout.Height(18)))
            {
                _activeCategories.Clear();
                _hiddenCategories.Clear();
            }
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();

            int btnPerRow = 6;
            for (int i = 0; i < Categories.Length; i++)
            {
                if (i % btnPerRow == 0) GUILayout.BeginHorizontal();

                var (label, type) = Categories[i];
                bool isShow = _activeCategories.Contains(type);
                bool isHide = _hiddenCategories.Contains(type);
                GUIStyle style = isShow ? _categoryBtnOn : (isHide ? _categoryBtnHide : _categoryBtnOff);

                bool catRightDown = Event.current.type == EventType.MouseDown && Event.current.button == 1;
                Vector2 catMousePos = Event.current.mousePosition;
                bool leftClicked = GUILayout.Button(new GUIContent(label, type), style, GUILayout.Width(32), GUILayout.Height(24));
                Rect btnRect = GUILayoutUtility.GetLastRect();

                if (leftClicked)
                {
                    if (isShow) _activeCategories.Remove(type);
                    else { _activeCategories.Add(type); _hiddenCategories.Remove(type); }
                }
                else if (catRightDown && btnRect.Contains(catMousePos))
                {
                    if (isHide) _hiddenCategories.Remove(type);
                    else { _hiddenCategories.Add(type); _activeCategories.Remove(type); }
                }

                if (i % btnPerRow == btnPerRow - 1 || i == Categories.Length - 1)
                    GUILayout.EndHorizontal();
            }

            // Show active filter state
            if (_activeCategories.Count > 0 || _hiddenCategories.Count > 0)
            {
                string parts = "";
                if (_activeCategories.Count > 0) parts += $"show: {string.Join(", ", _activeCategories)}";
                if (_hiddenCategories.Count  > 0) parts += (parts.Length > 0 ? "  |  " : "") + $"hide: {string.Join(", ", _hiddenCategories)}";
                GUILayout.Label($"<i>{parts}</i>", _smallLabelStyle);
            }

            GUILayout.Space(3);

            // Column headers — click to sort
            GUILayout.BeginHorizontal();
            if (GUILayout.Button(ColLabel(SortColumn.Name, "Name"), _colHeaderStyle, GUILayout.MinWidth(40f)))
                ClickSort(SortColumn.Name);
            if (GUILayout.Button(ColLabel(SortColumn.Type, "Type"), _colHeaderStyle, GUILayout.Width(50)))
                ClickSort(SortColumn.Type);
            if (GUILayout.Button(ColLabel(SortColumn.Distance, "km"), _colHeaderStyle, GUILayout.Width(44)))
                ClickSort(SortColumn.Distance);
            GUILayout.Label("Tgt",  _headerStyle, GUILayout.Width(24));
            GUILayout.Label("Ctrl", _headerStyle, GUILayout.Width(28));
            GUILayout.Label("M",    _headerStyle, GUILayout.Width(24));
            GUILayout.Space(HighLogic.Skin.verticalScrollbar.fixedWidth);
            GUILayout.EndHorizontal();

            // Scrollable rows — always show vertical scrollbar so its width is constant,
            // keeping the header labels above aligned with the data columns below.
            _scrollPos = GUILayout.BeginScrollView(_scrollPos, false, true, GUIStyle.none,
                HighLogic.Skin.verticalScrollbar, GUILayout.ExpandHeight(true));

            string filter = _filterText.ToLowerInvariant();
            int rowIdx = 0;

            foreach (var entry in _entries)
            {
                if (_activeCategories.Count > 0 && !_activeCategories.Contains(entry.Type))
                    continue;
                if (_hiddenCategories.Contains(entry.Type))
                    continue;

                if (!string.IsNullOrEmpty(filter) &&
                    !entry.Name.ToLowerInvariant().Contains(filter) &&
                    !entry.Type.ToLowerInvariant().Contains(filter))
                    continue;

                GUIStyle rowStyle;
                if (entry.IsActive)        rowStyle = _activeStyle;
                else if (entry.IsTargeted) rowStyle = _targetedStyle;
                else                       rowStyle = rowIdx % 2 == 0 ? _rowStyleA : _rowStyleB;

                GUILayout.BeginHorizontal(rowStyle);

                // Expand arrow for vessels with docking ports
                bool hasVessel = entry.Vessel != null;
                bool expanded  = hasVessel && _expandedVessels.Contains(entry.Vessel);
                string arrow   = expanded ? "▼" : (hasVessel ? "▶" : " ");

                // Name + expand toggle (left-click only via Input check)
                string eid      = entry.Vessel?.id.ToString();
                bool isFav      = eid != null && _favouriteIds.Contains(eid);
                bool hasNotes   = eid != null && _vesselNotes.ContainsKey(eid);
                string nameLabel = entry.IsActive ? $"► {entry.Name}" : entry.Name;
                if (isFav)    nameLabel = "* " + nameLabel;
                if (hasNotes) nameLabel += " [N]";
                if (GUILayout.Button($"{arrow} {nameLabel}", GUI.skin.label, GUILayout.MinWidth(40f)))
                {
                    if (hasVessel)
                    {
                        if (expanded) _expandedVessels.Remove(entry.Vessel);
                        else          _expandedVessels.Add(entry.Vessel);
                    }
                    FocusOn(entry);
                }

                GUILayout.Label(entry.Type, GUILayout.Width(50));

                string distStr = entry.Distance > 0 ? (entry.Distance / 1000.0).ToString("N0") : "—";
                GUILayout.Label(distStr, GUILayout.Width(44));

                // Target button
                if (!entry.IsActive)
                {
                    string tLabel = entry.IsTargeted ? "X" : "T";
                    if (GUILayout.Button(tLabel, GUILayout.Width(24)))
                    {
                        var av = FlightGlobals.ActiveVessel;
                        if (av != null)
                        {
                            ITargetable newTarget = entry.IsTargeted ? null :
                                                    entry.Vessel != null ? (ITargetable)entry.Vessel :
                                                    (ITargetable)entry.Body;
                            _deferredAction = () => { av.targetObject = newTarget; RefreshEntries(); };
                        }
                    }
                }
                else
                {
                    GUILayout.Label("", GUILayout.Width(24));
                }

                // Control button — switch to and control this vessel
                if (hasVessel && !entry.IsActive)
                {
                    if (GUILayout.Button("Ctl", GUILayout.Width(28)))
                    {
                        var v = entry.Vessel;
                        _deferredAction = () => ControlVessel(v);
                    }
                }
                else
                {
                    GUILayout.Label("", GUILayout.Width(28));
                }

                // Options button — vessels only
                if (hasVessel)
                {
                    if (GUILayout.Button("M", GUILayout.Width(24)))
                        OpenContextMenu(entry);
                }
                else
                {
                    GUILayout.Label("", GUILayout.Width(24));
                }

                GUILayout.EndHorizontal();

                // ── Docking port sub-rows ──────────────────────────────────────
                if (expanded && hasVessel)
                {
                    var allPorts = GetPorts(entry.Vessel);
                    var ports = new List<PortInfo>();
                    foreach (var p in allPorts)
                        if (p.Compatible) ports.Add(p);

                    if (allPorts.Count == 0)
                    {
                        GUILayout.BeginHorizontal(_portRowStyle);
                        GUILayout.Label("   (no docking ports)", _smallLabelStyle);
                        GUILayout.EndHorizontal();
                    }
                    else if (ports.Count == 0)
                    {
                        GUILayout.BeginHorizontal(_portRowStyle);
                        GUILayout.Label("   (no compatible ports)", _smallLabelStyle);
                        GUILayout.EndHorizontal();
                    }
                    else
                    {
                        // Port header
                        GUILayout.BeginHorizontal();
                        GUILayout.Space(8);
                        GUILayout.Label("<b>Port</b>",    _subHeaderStyle, GUILayout.MinWidth(40f));
                        GUILayout.Label("<b>Ang°</b>",    _subHeaderStyle, GUILayout.Width(36));
                        GUILayout.Label("<b>Roll°</b>",   _subHeaderStyle, GUILayout.Width(36));
                        GUILayout.Label("<b>m</b>",       _subHeaderStyle, GUILayout.Width(44));
                        GUILayout.Label("<b>m/s</b>",     _subHeaderStyle, GUILayout.Width(32));
                        GUILayout.Label("<b>Tgt</b>",     _subHeaderStyle, GUILayout.Width(24));
                        GUILayout.EndHorizontal();

                        var currentTarget = FlightGlobals.ActiveVessel?.targetObject;

                        foreach (var port in ports)
                        {
                            bool portTargeted = currentTarget is ModuleDockingNode pn && pn == port.Node;
                            GUIStyle pStyle = port.Occupied ? _portOccupiedStyle : _portRowStyle;

                            GUILayout.BeginHorizontal(pStyle);
                            GUILayout.Space(8);

                            string occupiedMark = port.Occupied ? "OCC" : "OPN";
                            string portLabel = $"{occupiedMark} {port.PartName}";
                            GUILayout.Label(portLabel, _smallLabelStyle, GUILayout.MinWidth(40f));

                            // Alignment info
                            GUILayout.Label(port.Angle.ToString("F0"),        _smallLabelStyle, GUILayout.Width(36));
                            GUILayout.Label(port.Roll.ToString("F0"),         _smallLabelStyle, GUILayout.Width(36));
                            GUILayout.Label(port.Distance.ToString("F0"),     _smallLabelStyle, GUILayout.Width(44));
                            GUILayout.Label(port.ApproachSpeed.ToString("F1"), _smallLabelStyle, GUILayout.Width(32));

                            // Target port button (only available when vessel is loaded)
                            if (!port.Occupied && port.Node != null)
                            {
                                string ptLabel = portTargeted ? "X" : "T";
                                if (GUILayout.Button(ptLabel, GUILayout.Width(22)))
                                {
                                    var av = FlightGlobals.ActiveVessel;
                                    if (av != null)
                                    {
                                        av.targetObject = portTargeted ? null : (ITargetable)port.Node;
                                        RefreshEntries();
                                    }
                                }
                            }
                            else
                            {
                                GUILayout.Label("–", _smallLabelStyle, GUILayout.Width(22));
                            }

                            GUILayout.EndHorizontal();
                        }
                    }
                }

                rowIdx++;
            }

            GUILayout.EndScrollView();

            // ── Rendezvous panel ──────────────────────────────────────────────
            GUILayout.Space(4);
            GUILayout.BeginVertical(HighLogic.Skin.box);
            if (_rdv.Valid)
            {
                GUILayout.Label($"<b>Rendezvous:</b> {_rdv.TargetName}", _headerStyle);

                GUILayout.BeginHorizontal();
                GUILayout.Label($"Phase {_rdv.CurPhase:F0}° → {_rdv.ReqPhase:F0}°", _smallLabelStyle, GUILayout.Width(150));
                GUILayout.Label($"Win: {FormatDuration(_rdv.TimeToWindow)}", _smallLabelStyle);
                GUILayout.EndHorizontal();

                GUILayout.BeginHorizontal();
                GUILayout.Label($"Hohm: {_rdv.HohmannDv:F0} m/s", _smallLabelStyle, GUILayout.Width(150));
                GUILayout.Label($"Xfr: {FormatDuration(_rdv.TransferTime)}", _smallLabelStyle);
                GUILayout.EndHorizontal();

                if (_rdv.RelInclination > 0.05)
                {
                    GUILayout.BeginHorizontal();
                    GUILayout.Label($"Inc: {_rdv.RelInclination:F1}°", _smallLabelStyle, GUILayout.Width(150));
                    GUILayout.Label($"Pln: {_rdv.PlaneChangeDv:F0} m/s", _smallLabelStyle);
                    GUILayout.EndHorizontal();
                }
            }
            else
            {
                GUILayout.Label("<b>Rendezvous:</b> target a vessel", _headerStyle);
            }
            GUILayout.EndVertical();

            GUILayout.Space(3);
            if (GUILayout.Button("Refresh Now"))
                RefreshEntries();

            GUILayout.Space(2);
            GUILayout.EndVertical();

            // Resize grip — drawn at bottom-right in window-local coords
            GUI.Label(new Rect(_windowRect.width - 18, _windowRect.height - 18, 16, 16), "◢", _resizeGripStyle);

            GUI.DragWindow(new Rect(0, 0, 10000, 24));
        }

        // ── Focus + Control Logic ──────────────────────────────────────────────

        private void FocusOn(VesselEntry entry)
        {
            if (PlanetariumCamera.fetch != null)
            {
                MapObject mo = entry.Vessel?.mapObject ?? entry.Body?.MapObject;
                if (mo != null) PlanetariumCamera.fetch.SetTarget(mo);
            }
            RefreshEntries();
        }

        private void ControlVessel(Vessel v)
        {
            if (v == null || v == FlightGlobals.ActiveVessel) return;
            FlightGlobals.SetActiveVessel(v);

            // Put the camera on the new vessel too
            if (PlanetariumCamera.fetch != null && v.mapObject != null)
                PlanetariumCamera.fetch.SetTarget(v.mapObject);

            RefreshEntries();
        }

        // ── Context Menu ───────────────────────────────────────────────────────

        private void OpenContextMenu(VesselEntry entry)
        {
            _contextMenuEntry = entry;
            _contextMenuOpen  = true;
            float mx = Input.mousePosition.x;
            float my = Screen.height - Input.mousePosition.y;
            bool hasVessel = entry.Vessel != null;
            float menuHeight = hasVessel ? (!entry.IsActive ? 140f : 100f) : 55f;
            _contextMenuRect = new Rect(
                Mathf.Min(mx, Screen.width  - 165f),
                Mathf.Min(my, Screen.height - menuHeight),
                160f, menuHeight
            );
        }

        private void DrawContextMenu(int id)
        {
            var    entry    = _contextMenuEntry;
            string eid      = entry.Vessel?.id.ToString();
            bool   isFav    = eid != null && _favouriteIds.Contains(eid);
            bool   hasNotes = eid != null && _vesselNotes.ContainsKey(eid);

            GUILayout.BeginVertical();

            if (entry.Vessel != null)
            {
                string noteLabel = hasNotes ? "Show Notes [N]" : "Show Notes";
                if (GUILayout.Button(noteLabel, _popupBtnStyle))
                {
                    _contextMenuOpen = false;
                    OpenNotesWindow(entry);
                }

                if (GUILayout.Button(isFav ? "Unfavourite" : "Favourite", _popupBtnStyle))
                {
                    if (isFav) _favouriteIds.Remove(eid);
                    else       _favouriteIds.Add(eid);
                    ApplySort();
                    _contextMenuOpen = false;
                }
            }

            if (GUILayout.Button("Show on Map", _popupBtnStyle))
            {
                ShowOnMap(entry);
                _contextMenuOpen = false;
            }

            if (entry.Vessel != null)
            {
                if (GUILayout.Button("Rename", _popupBtnStyle))
                {
                    _contextMenuOpen = false;
                    OpenRenameWindow(entry);
                }

                if (!entry.IsActive && GUILayout.Button("Terminate Vessel", _dangerBtnStyle))
                {
                    _contextMenuOpen = false;
                    OpenTerminateWindow(entry);
                }
            }

            GUILayout.EndVertical();
            GUI.DragWindow(new Rect(0, 0, 10000, 24));
        }

        private void ShowOnMap(VesselEntry entry)
        {
            MapView.EnterMapView();
            if (PlanetariumCamera.fetch != null)
            {
                MapObject mo = entry.Vessel?.mapObject ?? entry.Body?.MapObject;
                if (mo != null) PlanetariumCamera.fetch.SetTarget(mo);
            }
        }

        // ── Notes Popup ────────────────────────────────────────────────────────

        private void OpenNotesWindow(VesselEntry entry)
        {
            _notesOpen     = true;
            _notesEntry    = entry;
            _notesId       = entry.Vessel.id.ToString();
            _notesEditText = _vesselNotes.TryGetValue(_notesId, out string existing) ? existing : "";
            _notesRect     = new Rect(Screen.width / 2f - 160f, Screen.height / 2f - 120f, 320f, 240f);
        }

        private void DrawNotesWindow(int id)
        {
            GUILayout.BeginVertical();
            GUILayout.Label($"<b>{_notesEntry.Name}</b>", _headerStyle);
            _notesEditText = GUILayout.TextArea(_notesEditText, GUILayout.ExpandHeight(true));
            GUILayout.BeginHorizontal();
            if (GUILayout.Button("Save"))
            {
                if (string.IsNullOrWhiteSpace(_notesEditText))
                    _vesselNotes.Remove(_notesId);
                else
                    _vesselNotes[_notesId] = _notesEditText;
                SaveNotes();
                _notesOpen = false;
            }
            if (GUILayout.Button("Cancel"))
                _notesOpen = false;
            GUILayout.EndHorizontal();
            GUILayout.EndVertical();
            GUI.DragWindow(new Rect(0, 0, 10000, 20));
        }

        // ── Rename Popup ───────────────────────────────────────────────────────

        private void OpenRenameWindow(VesselEntry entry)
        {
            _renameOpen  = true;
            _renameEntry = entry;
            _renameText  = entry.Name;
            _renameRect  = new Rect(Screen.width / 2f - 150f, Screen.height / 2f - 50f, 300f, 100f);
        }

        private void DrawRenameWindow(int id)
        {
            GUILayout.BeginVertical();
            GUILayout.Label($"Renaming: <b>{_renameEntry.Name}</b>", _headerStyle);
            GUI.SetNextControlName("renameField");
            _renameText = GUILayout.TextField(_renameText);
            if (Event.current.type == EventType.KeyDown && Event.current.keyCode == KeyCode.Return)
            {
                ApplyRename();
                Event.current.Use();
            }
            GUILayout.BeginHorizontal();
            if (GUILayout.Button("Rename")) ApplyRename();
            if (GUILayout.Button("Cancel")) _renameOpen = false;
            GUILayout.EndHorizontal();
            GUILayout.EndVertical();
            GUI.DragWindow(new Rect(0, 0, 10000, 20));
        }

        private void ApplyRename()
        {
            if (!string.IsNullOrWhiteSpace(_renameText) && _renameEntry.Vessel != null)
            {
                _renameEntry.Vessel.vesselName = _renameText.Trim();
                RefreshEntries();
            }
            _renameOpen = false;
        }

        // ── Terminate Popup ────────────────────────────────────────────────────

        private void OpenTerminateWindow(VesselEntry entry)
        {
            _terminateOpen  = true;
            _terminateEntry = entry;
            _terminateRect  = new Rect(Screen.width / 2f - 125f, Screen.height / 2f - 50f, 250f, 100f);
        }

        private void DrawTerminateWindow(int id)
        {
            GUILayout.BeginVertical();
            GUILayout.Label($"Terminate <b>{_terminateEntry.Name}</b>?", _headerStyle);
            GUILayout.Label("This action cannot be undone.", _smallLabelStyle);
            GUILayout.Space(4f);
            GUILayout.BeginHorizontal();
            if (GUILayout.Button("Terminate", _dangerBtnStyle))
            {
                _terminateEntry.Vessel?.Die();
                _terminateOpen = false;
                RefreshEntries();
            }
            if (GUILayout.Button("Cancel"))
                _terminateOpen = false;
            GUILayout.EndHorizontal();
            GUILayout.EndVertical();
            GUI.DragWindow(new Rect(0, 0, 10000, 20));
        }

        // ── Persistence ────────────────────────────────────────────────────────

        private void LoadData()
        {
            try
            {
                var cfg = KSP.IO.PluginConfiguration.CreateForType<FocusTableMod>();
                cfg.load();
                _vesselNotes.Clear();
                int count = cfg.GetValue("noteCount", 0);
                for (int i = 0; i < count; i++)
                {
                    string id   = cfg.GetValue("noteId"   + i, "");
                    string text = cfg.GetValue("noteText" + i, "");
                    if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(text))
                        _vesselNotes[id] = text;
                }
            }
            catch (Exception e)
            {
                Debug.LogWarning("[FocusTable] LoadData failed: " + e.Message);
            }
        }

        private void SaveNotes()
        {
            try
            {
                var cfg = KSP.IO.PluginConfiguration.CreateForType<FocusTableMod>();
                cfg.load();
                cfg.SetValue("noteCount", _vesselNotes.Count);
                int i = 0;
                foreach (var kv in _vesselNotes)
                {
                    cfg.SetValue("noteId"   + i, kv.Key);
                    cfg.SetValue("noteText" + i, kv.Value);
                    i++;
                }
                cfg.save();
            }
            catch (Exception e)
            {
                Debug.LogWarning("[FocusTable] SaveNotes failed: " + e.Message);
            }
        }

        // ── Styles ─────────────────────────────────────────────────────────────

        private void InitStyles()
        {
            if (_stylesInitialized) return;

            _headerStyle = new GUIStyle(HighLogic.Skin.label)
            {
                fontSize = 11, richText = true, wordWrap = false,
                normal = { textColor = new Color(0.85f, 0.85f, 0.85f) }
            };

            _subHeaderStyle = new GUIStyle(HighLogic.Skin.label)
            {
                fontSize = 10, richText = true, wordWrap = false,
                normal = { textColor = new Color(0.6f, 0.8f, 1f) }
            };

            _smallLabelStyle = new GUIStyle(HighLogic.Skin.label)
            {
                fontSize = 10, wordWrap = false,
                normal = { textColor = new Color(0.75f, 0.75f, 0.75f) }
            };

            _rowStyleA = new GUIStyle(GUI.skin.box)
            {
                padding = new RectOffset(2, 2, 2, 2), margin = new RectOffset(0, 0, 1, 0),
                normal  = { background = MakeTex(1, 1, new Color(0.18f, 0.18f, 0.22f, 0.85f)) }
            };

            _rowStyleB = new GUIStyle(_rowStyleA)
            {
                normal = { background = MakeTex(1, 1, new Color(0.13f, 0.13f, 0.17f, 0.85f)) }
            };

            _activeStyle = new GUIStyle(_rowStyleA)
            {
                normal = { background = MakeTex(1, 1, new Color(0.15f, 0.35f, 0.55f, 0.95f)) }
            };

            _targetedStyle = new GUIStyle(_rowStyleA)
            {
                normal = { background = MakeTex(1, 1, new Color(0.35f, 0.20f, 0.50f, 0.95f)) }
            };

            _portRowStyle = new GUIStyle(_rowStyleA)
            {
                normal = { background = MakeTex(1, 1, new Color(0.10f, 0.20f, 0.15f, 0.90f)) }
            };

            _portOccupiedStyle = new GUIStyle(_rowStyleA)
            {
                normal = { background = MakeTex(1, 1, new Color(0.25f, 0.15f, 0.10f, 0.90f)) }
            };

            _filterStyle = new GUIStyle(HighLogic.Skin.textField) { fontSize = 11 };

            _categoryBtnOff = new GUIStyle(HighLogic.Skin.button)
            {
                fontSize = 13, padding = new RectOffset(2, 2, 1, 1),
                normal   = { textColor = new Color(0.75f, 0.75f, 0.75f) }
            };

            _categoryBtnOn = new GUIStyle(_categoryBtnOff)
            {
                normal  = { background = MakeTex(1, 1, new Color(0.2f, 0.5f, 0.8f, 1f)), textColor = Color.white },
                focused = { background = MakeTex(1, 1, new Color(0.2f, 0.5f, 0.8f, 1f)), textColor = Color.white }
            };

            _categoryBtnHide = new GUIStyle(_categoryBtnOff)
            {
                normal  = { background = MakeTex(1, 1, new Color(0.6f, 0.15f, 0.15f, 1f)), textColor = Color.white },
                focused = { background = MakeTex(1, 1, new Color(0.6f, 0.15f, 0.15f, 1f)), textColor = Color.white }
            };

            _colHeaderStyle = new GUIStyle(_headerStyle)
            {
                normal  = { textColor = new Color(0.85f, 0.85f, 0.85f), background = null },
                hover   = { textColor = new Color(0.4f,  0.8f,  1.0f),  background = null },
                active  = { textColor = new Color(0.6f,  0.9f,  1.0f),  background = null },
                focused = { textColor = new Color(0.85f, 0.85f, 0.85f), background = null },
                richText = true
            };

            _resizeGripStyle = new GUIStyle(HighLogic.Skin.label)
            {
                fontSize = 14,
                normal   = { textColor = new Color(0.55f, 0.55f, 0.55f, 0.8f) }
            };

            _popupBtnStyle = new GUIStyle(HighLogic.Skin.button)
            {
                fontSize = 11,
                padding  = new RectOffset(8, 8, 4, 4)
            };

            _dangerBtnStyle = new GUIStyle(_popupBtnStyle)
            {
                normal = { background = MakeTex(1, 1, new Color(0.55f, 0.10f, 0.10f, 1f)), textColor = Color.white },
                hover  = { background = MakeTex(1, 1, new Color(0.75f, 0.15f, 0.15f, 1f)), textColor = Color.white },
                active = { background = MakeTex(1, 1, new Color(0.45f, 0.08f, 0.08f, 1f)), textColor = Color.white }
            };

            _stylesInitialized = true;
        }

        private static Texture2D MakeTex(int w, int h, Color col)
        {
            var tex = new Texture2D(w, h, TextureFormat.ARGB32, false);
            var pixels = new Color[w * h];
            for (int i = 0; i < pixels.Length; i++) pixels[i] = col;
            tex.SetPixels(pixels);
            tex.Apply();
            return tex;
        }
    }
}
