Attribute VB_Name = "modCodejock17"
Public frmInbox As frmInbox
Public frmShortBar As frmShortcutBar2
''Public frmPreviewPane As frmPreviewPane
Public frmEditEvent As frmEditEvent
'Public frmPaneMail As frmPaneMail
'Public frmPaneJournal As frmPaneJournal
'Public frmPaneTasks As frmPaneTasks
Public frmPaneContacts As frmPaneContacts2
Public frmPaneCalendar As frmPaneCalendar2
'Public frmPaneShortcuts As frmPaneShortcuts
'Public frmPaneNotes As frmPaneNotes




'Public pageBackstageInfo As pageBackstageInfo
Public pageBackstageHelp As pageBackstageHelp
'Public pageBackstageSend As pageBackstageSend

Public Const ID_SWITCH_PRINTLAYOUT = 7700
Public Const ID_SWITCH_FULLSCREENREADING = 7701
Public Const ID_SWITCH_WEBLAYOUT = 7702
Public Const ID_SWITCH_OUTLINE = 7703
Public Const ID_SWITCH_DRAFT = 7704

Public Const ID_SWITCH_NORMAL = 7705
Public Const ID_SWITCH_CALENAR_AND_TASK = 7706
Public Const ID_SWITCH_CALENDAR = 7707
Public Const ID_SWITCH_CLASSIC = 7708
Public Const ID_SWITCH_READING = 7709

Public Const ID_GALLERY_QUICKSTEP = 7750
Public Const ID_QUICKSTEP_MOVE_TO = 7755
Public Const ID_QUICKSTEP_TEAM_EMAIL = 7756
Public Const ID_QUICKSTEP_REPLAY_DELETE = 7757
Public Const ID_QUICKSTEP_TO_MANAGER = 7758
Public Const ID_QUICKSTEP_DONE = 7759
Public Const ID_QUICKSTEP_CREATE_NEW = 7760
Public Const ID_GROUP_QUICKSTEP = 7761

Public Const ID_QUICKSTEP_CATEGORIZE = 7762
Public Const ID_QUICKSTEP_FLAG_MOVE = 7763

Public Const ID_GROUP_MAIL_NEW = 7764
Public Const ID_GROUP_MAIL_NEW_NEW = 7765
Public Const ID_GROUP_MAIL_NEW_NEW_ITEMS = 7766
Public Const ID_GROUP_MAIL_NEW_APPLOINTMENT = 7767
Public Const ID_GROUP_MAIL_NEW_CONTACT = 7768
Public Const ID_GROUP_MAIL_NEW_TASK = 7769
Public Const ID_GROUP_MAIL_DELETE = 7770
Public Const ID_GROUP_MAIL_DELETE_CLEANUP = 7771
Public Const ID_GROUP_MAIL_DELETE_JUNK = 7772
Public Const ID_GROUP_MAIL_DELETE_DELETE = 7773
Public Const ID_GROUP_MAIL_RESPOND = 7774
Public Const ID_GROUP_MAIL_RESPOND_REPLY = 7775
Public Const ID_GROUP_MAIL_RESPOND_REPLY_ALL = 7776
Public Const ID_GROUP_MAIL_RESPOND_FORWARD = 7777
Public Const ID_GROUP_MAIL_RESPOND_MEETING = 7778
Public Const ID_GROUP_MAIL_RESPOND_IM = 7779
Public Const ID_GROUP_MAIL_RESPOND_MORE = 7780
Public Const ID_GROUP_MAIL_MOVE = 7781
Public Const ID_GROUP_MAIL_MOVE_MOVE = 7782
Public Const ID_GROUP_MAIL_MOVE_ONENOTE = 7783
Public Const ID_GROUP_MAIL_TAGS = 7784
Public Const ID_GROUP_MAIL_TAGS_UNREAD = 7785
Public Const ID_GROUP_MAIL_TAGS_CATEGORIZE = 7786
Public Const ID_GROUP_MAIL_TAGS_FOLLOWUP = 7787
Public Const ID_GROUP_MAIL_FIND = 7788
Public Const ID_GROUP_MAIL_FIND_CONTACT = 7789
Public Const ID_GROUP_MAIL_FIND_ADDRESSBOOK = 7790
Public Const ID_GROUP_MAIL_FIND_FILTER = 7791

Public Const ID_INDICATOR_PAGENUMBER = 220
Public Const ID_INDICATOR_WORDCOUNT = 221
Public Const ID_INDICATOR_LANGUAGE = 222
Public Const ID_INDICATOR_TRACKCHANGES = 223
Public Const ID_INDICATOR_CAPSLOCK = 224
Public Const ID_INDICATOR_OVERTYPE = 225
Public Const ID_INDICATOR_MACRORECORDING = 226
Public Const ID_INDICATOR_VIEWSHORTCUTS = 227
Public Const ID_INDICATOR_ZOOM = 228
Public Const ID_INDICATOR_ZOOMSLIDER = 229

Public Const ID_OPTIONS_RTL = 3004
Public Const ID_OPTIONS_ANIMATION = 3005

Public Const ID_OPTIONS_STYLEBLUE2007 = 2890
Public Const ID_OPTIONS_STYLESILVER2007 = 2891
Public Const ID_OPTIONS_STYLEBLACK2007 = 2892
Public Const ID_OPTIONS_STYLEAQUA2007 = 2893
Public Const ID_OPTIONS_STYLESCENIC7 = 2894
Public Const ID_OPTIONS_STYLEBLUE2010 = 2896
Public Const ID_OPTIONS_STYLESILVER2010 = 2897
Public Const ID_OPTIONS_STYLEBLACK2010 = 2898
Public Const ID_OPTIONS_STYLESYSTEM = 2899

Public Const ID_RIBBON_MINIMIZE = 4567
Public Const ID_RIBBON_EXPAND = 4568
Public Const ID_RIBBON_QUICKACCESSEMPTYICON = 302

Public Const IDR_CNTR_INPLACE = 6
Public Const IDD_ABOUTBOX = 100
Public Const IDP_OLE_INIT_FAILED = 100
Public Const IDP_FAILED_TO_CREATE = 102
Public Const IDR_MAINFRAME = 128
Public Const IDR_SMALLICONS = 128
Public Const IDR_RIBBONTYPE = 129
Public Const IDR_LARGEICONS = 131
Public Const IDR_LAYOUTTABSMALL = 143
Public Const IDR_LAYOUTTABLARGE = 145
Public Const IDR_MENU_CONTEXT = 147
Public Const ID_APP_THEME = 148
Public Const IDB_BITMAP_PICTURE = 149
Public Const IDB_BITMAP_GRAPHIC = 150
Public Const IDB_BITMAP_CHART = 151
Public Const IDB_BITMAP_TABLE = 152
Public Const IDB_INSERTTAB = 200
Public Const IDB_WRITETAB = 201
Public Const IDB_BITMAPS_GROUPS = 202
Public Const IDB_GEAR = 300
Public Const ID_GROUP_BUTTONPOPUP = 2000
Public Const ID_INSERT_HYPERLINK = 2710
Public Const ID_INSERT_CROSS_REFERENCE = 2712
Public Const ID_TEXT_SIGNATURE = 2713
Public Const ID_TEXT_DATETIME = 2714
Public Const ID_TEXT_INSERTOBJECT = 2715
Public Const ID_CANCEL_EDIT_CNTR = 2768
Public Const ID_PAGES_NEW = 2772
Public Const ID_PAGES_COVRE = 2773
Public Const ID_PAGES_BREAK = 2774
Public Const ID_TABLE_NEW = 2775
Public Const ID_ILLUSTRATION_PICTURE = 2776
Public Const ID_ILLUSTRATION_GRAPHIC = 2777
Public Const ID_ILLUSTRATION_CHART = 2778
Public Const ID_TABLE_INSERTTABLE = 2779
Public Const ID_INSERT_HEADER = 2780
Public Const ID_INSERT_FOOTER = 2781
Public Const ID_INSERT_PAGENUMBER = 2782
Public Const ID_TEXT_TEXTBOX = 2783
Public Const ID_TEXT_PARTS = 2784
Public Const ID_TEXT_WORDART = 2785
Public Const ID_TEXT_DROPCAP = 2786
Public Const ID_SYMBOL_EQUATIONS = 2787
Public Const ID_SYMBOL_SYMBOL = 2788
Public Const ID_PAGES_COVER = 2789
Public Const ID_ILLUSTRATION_CLIPART = 2790
Public Const ID_ILLUSTRATION_FROMCAMERA = 2791
Public Const ID_INSERT_BOOKMARK = 2791
Public Const ID_PAGENUMBER_FORMATPAGENUMBERS = 2792
Public Const ID_FONT_GROW = 2792
Public Const ID_PAGENUMBER_REMOVEPAGENUMBERS = 2793
Public Const ID_FONT_SHRINK = 2793
Public Const ID_NEWPAGE_BLANKPAGE = 2794
Public Const ID_FONT_CLEAR = 2794
Public Const ID_NEWPAGE_SELECTION = 2795
Public Const ID_TEXT_CHANGECASE = 2795
Public Const ID_INSERT_NUMBERING = 2796
Public Const ID_INSERT_LIST = 2797
Public Const ID_PARA_DECREASEINDENT = 2798
Public Const ID_PARA_INCREASEINDENT = 2799
Public Const ID_PARA_SORT = 2800
Public Const ID_PARA_JUSTIFY = 2801
Public Const ID_PARA_SHOWMARKS = 2802
Public Const ID_DOCUMENTPARTS_AUTOTEXT = 2803
Public Const ID_PARA_LINESPACING = 2803
Public Const ID_DOCUMENTPARTS_PROPERTY = 2804
Public Const ID_PARA_SHADING = 2804
Public Const ID_DOCUMENTPARTS_FIELD = 2805
Public Const ID_BORDERS_NOBORDER = 2805
Public Const ID_DOCUMENTPARTS_BUILDINGBLOCKORGANIZER = 3806
Public Const ID_TEXT_HIGHLIGHTCOLOR = 2806
Public Const ID_VIEW_RULER = 2807
Public Const ID_VIEW_GRIDLINES = 2808
Public Const ID_VIEW_PROPERTIES = 2809
Public Const ID_VIEW_DOCUMENTMAP = 2810
Public Const ID_VIEW_THUMBNAILS = 2811
Public Const ID_VIEW_ACTINBAR = 2812
Public Const ID_TEXT_COLOR_SELECTOR = 2813
Public Const ID_GROUP_PARAGRAPH = 5000
Public Const ID_GROUP_CLIPBOARD = 5001
Public Const ID_GROUP_FONT = 5002
Public Const ID_GROUP_FIND = 5003
Public Const ID_TAB_WRITE = 5004
Public Const ID_TAB_INSERT = 5005
Public Const ID_TAB_PAGELAYOUT = 5006
Public Const ID_TAB_ADDINS = 5007
Public Const ID_TAB_TABLEDESIGN = 5008
Public Const ID_TAB_TABLELAYOUT = 5009
Public Const ID_TAB_CHARTDESIGN = 5010
Public Const ID_TAB_CHARTFORMAT = 5011
Public Const ID_TAB_CHARTLAYOUT = 5012
Public Const ID_TAB_CONTEXTCHART = 5013
Public Const ID_GROUP_PAGES = 5014
Public Const ID_GROUP_TABLE = 5015
Public Const ID_GROUP_ILLUSTRATIONS = 5016
Public Const ID_GROUP_HEADERFOOTERS = 5017
Public Const ID_GROUP_LINKS = 5018
Public Const ID_GROUP_TEXT = 5019
Public Const ID_GROUP_SYMBOLS = 5020
Public Const ID_GROUP_THEMES = 5021
Public Const ID_GROUP_PAGESETUP = 5022
Public Const ID_GROUP_PAGEBACKGROUND = 5023
Public Const ID_GROUP_ARRANGE = 5024
Public Const ID_GROUP_SHOWHIDE = 5025
Public Const ID_TAB_VIEW = 5026
Public Const ID_TAB_REFERENCES = 5027
Public Const ID_TAB_MAILINGS = 5028
Public Const ID_TAB_REVIEW = 5029
Public Const ID_CHAR_BOLD = 7608
Public Const ID_CHAR_ITALIC = 7610
Public Const ID_CHAR_UNDERLINE = 7611
Public Const ID_EDIT_GOTO = 7612
Public Const ID_EDIT_SELECT_OBJECTS = 7613
Public Const ID_EDIT_SELECT = 7614
Public Const ID_EDIT_SELECT_MULTIPLE_OBJECTS = 7615
Public Const ID_FORMAT_PAINTER = 7616
Public Const ID_TEXT_FONT = 7617
Public Const ID_FONT_FACE = 7618
Public Const ID_FONT_SIZE = 7619
Public Const ID_CHAR_STRIKETHROUGH = 7620
Public Const ID_TEXT_SUBSCRIPT = 7621
Public Const ID_TEXT_SUPERSCRIPT = 7622
Public Const ID_TEXT_COLOR = 7623
Public Const ID_INSERT_BULLET = 32777
Public Const ID_COVERPAGE_REMOVECURRENTCOVERPAGE = 32796
Public Const ID_COVERPAGE_SAVESELECTIONASNEWCOVERPAGE = 32797
Public Const ID_TABLE_DRAWTABLE = 32799
Public Const ID_TABLE_CONVERTTEXTTOTABLE = 32800
Public Const ID_TEXTBOX_DRAWTEXTBOX = 32801
Public Const ID_TEXTBOX_SAVESELECTIONASNEWTEXTBOX = 32802
Public Const ID_PARA_LEFT = 32803
Public Const ID_PARA_CENTER = 32804
Public Const ID_PARA_RIGHT = 32805
Public Const ID_EQUATIONS_MATH = 32807
Public Const ID_THEMES_COLORS = 32808
Public Const ID_THEMES_FONTS = 32809
Public Const ID_THEMES_EFFECTS = 32810
Public Const ID_PAGE_ORIENTATIONS = 32811
Public Const ID_PAGE_ORIENTATION = 32812
Public Const ID_PAGE_SIZE = 32813
Public Const ID_PAGE_COLUMNS = 32814
Public Const ID_PAGE_BREAKS = 32815
Public Const ID_PAGE_LINENUMBERS = 32816
Public Const ID_PAGE_HYPHENATATION = 32817
Public Const ID_PAGE_WATERMARK = 32818
Public Const ID_PAGE_COLOR = 32819
Public Const ID_PAGE_BORDERS = 32820
Public Const ID_ARRANGE_FRONT = 32821
Public Const ID_ARRANGE_BACK = 32822
Public Const ID_ARRANGE_ALIGN = 32823
Public Const ID_ARRANGE_GROUP = 32824
Public Const ID_ARRANGE_UNGROUP = 32825
Public Const ID_ARRANGE_ROTATE = 32826
Public Const ID_THEMES_THEMES = 32827
Public Const ID_PAGE_MARGINS = 32828
Public Const ID_ARRANGE_POSITION = 32829
Public Const ID_ARRANGE_TEXTWRAPPING = 32831
Public Const ID_CONTEXT_FONT = 32832
Public Const ID_PARA_PARAGRAPH = 32833
Public Const ID_THEME_OFFICE2003 = 32834
Public Const ID_THEME_OFFICE2007 = 32835
Public Const IDB_CLIENT_FACE = 3010
Public Const ID_SYSTEM_ICON = 1200
Public Const ID_FILE_PREPARE = 1230
Public Const ID_FILE_SEND_MAIL = 1231
Public Const ID_FILE_PUBLISH = 1232
Public Const ID_FILE_CLOSE = 1233
Public Const ID_FILE_SEND_INTERNETFAX = 1234
Public Const ID_FILE_SEND = 1235
Public Const ID_FILE_OPTIONS = 1236

Public Const ID_OPTIONS_FONT_SYSTEM = 42883
Public Const ID_OPTIONS_FONT_NORMAL = 42884
Public Const ID_OPTIONS_FONT_LARGE = 42885
Public Const ID_OPTIONS_FONT_EXTRALARGE = 42886
Public Const ID_OPTIONS_FONT_AUTORESIZEICONS = 42887


Public Const ID_GROUP_HEADERANDFOOTER = 2003
Public Const ID_GROUP_POPUPICON = 2004
Public Const ID_GROUP_STYLES = 2005
Public Const ID_GALLERY_STYLES = 2006
Public Const ID_GALLERY_SHAPES = 2007
Public Const ID_GALLERY_COLORS = 2010
Public Const ID_GALLERY_LARGE_COLORS_POPUP = 2012
Public Const ID_GROUP_SHAPES = 2205

Public Const ID_APP_ABOUT = 4000
Public Const ID_EDIT_PASTE = 4001
Public Const ID_EDIT_PASTE_SPECIAL = 4002
Public Const ID_EDIT_COPY = 4003
Public Const ID_EDIT_CUT = 4004
Public Const ID_EDIT_FIND = 57636
Public Const ID_EDIT_REPLACE = 4006
Public Const ID_EDIT_SELECT_ALL = 4007
Public Const ID_FILE_NEW = 4008
Public Const ID_FILE_OPEN = 4009
Public Const ID_FILE_SAVE = 4010
Public Const ID_FILE_PRINT = 4011
Public Const ID_FILE_SAVE_AS = 57604
Public Const ID_FILE_PRINT_PREVIEW = 57609
Public Const ID_FILE_PRINT_SETUP = 57606
Public Const ID_FILE_MRU_FILE1 = 57616
Public Const ID_APP_EXIT = 57665

Public Const ID_SEARCH_ICON = 57783

Public Const ID_SAMPLE_MENU_ITEM = 60006

Public Const ID_GROUP_CLIPBOARD_OPTION = 3400
Public Const ID_GROUP_FONT_OPTION = 3401


Public Const ID_OPTIONS_STYLEBLUE = 3000
Public Const ID_OPTIONS_STYLEBLACK = 3001
Public Const ID_OPTIONS_STYLEAQUA = 3002
Public Const ID_OPTIONS_STYLESILVER = 3003

Public Const ID_PARAGRAPH_INDENTLEFT = 4500
Public Const ID_PARAGRAPH_INDENTRIGHT = 4501
Public Const ID_PARAGRAPH_SPACINGBEFORE = 4502
Public Const ID_PARAGRAPH_SPACINGAFTER = 4503


'Commandbars public constants






Public Const ID_FILE_EXIT2 = 10004






'Public Const ID_EDIT_MOVE_TO_FOLDER = 110
'Public Const ID_EDIT_MARK_AS_READ = 111
'Public Const ID_EDIT_MARK_AS_UNREAD = 112
'Public Const ID_EDIT_DELETE = 113
'
'Public Const ID_VIEW_TOOLBAR_STANDARD = 114
'Public Const ID_VIEW_TOOLBAR_THEMES = 115
Public Const ID_VIEW_STATUSBAR = 10016
'Public Const ID_VIEW_OPTIONS = 117
'Public Const ID_VIEW_REFRESH = 123
'Public Const ID_VIEW_READING_PANE = 124
'Public Const ID_VIEW_READING_PANE_RIGHT = 216
'Public Const ID_VIEW_READING_PANE_BOTTOM = 217
'Public Const ID_VIEW_READING_PANE_OFF = 218
'
'Public Const ID_HELP_ABOUT = 118
'
'Public Const ID_TOOLS_SEND = 119
'Public Const ID_TOOLS_RECEIVE = 120
'Public Const ID_TOOLS_FIND = 121
'Public Const ID_TOOLS_ADDRESS_BOOK = 145
'Public Const ID_TOOLS_RULES_AND_ALERTS = 204
'
'Public Const ID_GO_MAIL = 130
'Public Const ID_GO_CALENDAR = 131
'Public Const ID_GO_CONTACTS = 132
'Public Const ID_GO_TASKS = 133
'Public Const ID_GO_NOTES = 134
'Public Const ID_GO_FOLDER_LIST = 135
'Public Const ID_GO_SHORTCUTS = 136
'Public Const ID_GO_JOURNAL = 137
'
'Public Const ID_ACTIONS_REPLY = 140
'Public Const ID_ACTIONS_REPLY_TO_ALL = 141
'Public Const ID_ACTIONS_FORWARD = 142
'Public Const ID_ACTIONS_SEND_RECEIVE = 143
'Public Const ID_ACTIONS_CREATE_RULE = 144
'
'Public Const ID_WEB_BACK = 146
'Public Const ID_WEB_FORWARD = 147
'Public Const ID_WEB_HOME = 148
'Public Const ID_WEB_SEARCH = 149
'Public Const ID_WEB_REFRESH = 200
'Public Const ID_WEB_STOP = 201
'
'Public Const ID_OUTLOOK_TODAY = 202
'Public Const ID_BACK = 210
'Public Const ID_FORWARD = 211
'
'Public Const ID_NAVIGATE_UP = 203
'Public Const ID_ALIGN_LEFT = 213
'Public Const ID_ALIGN_CENTER = 214
'Public Const ID_ALIGN_RIGHT = 215
'
'Public Const XTP_ID_TOOLBARLIST = 59392
'
'Public Const ID_REPORTCONTROL_ALLOWCOLUMNREMOVE = 150
'Public Const ID_REPORTCONTROL_ALLOWCOLUMNREORDER = 151
'Public Const ID_REPORTCONTROL_ALLOWCOLUMNRESIZE = 152
'Public Const ID_REPORTCONTROL_AUTOMATICALLYGROUPITEMS = 153
'Public Const ID_REPORTCONTROL_HORIZONTALGRIDSTYLE_DASHES = 154
'Public Const ID_REPORTCONTROL_HORIZONTALGRIDSTYLE_LARGEDOTS = 155
'Public Const ID_REPORTCONTROL_HORIZONTALGRIDSTYLE_SMALLDOTS = 156
'Public Const ID_REPORTCONTROL_HORIZONTALGRIDSTYLE_NOLINES = 157
'Public Const ID_REPORTCONTROL_HORIZONTALGRIDSTYLE_SOLID = 158
'Public Const ID_REPORTCONTROL_VERTICALGRIDSTYLE_DASHES = 159
'Public Const ID_REPORTCONTROL_VERTICALGRIDSTYLE_LARGEDOTS = 160
'Public Const ID_REPORTCONTROL_VERTICALGRIDSTYLE_SMALLDOTS = 161
'Public Const ID_REPORTCONTROL_VERTICALGRIDSTYLE_NOLINES = 162
'Public Const ID_REPORTCONTROL_VERTICALGRIDSTYLE_SOLID = 163
'Public Const ID_REPORTCONTROL_MULTIPLESELECTION = 164
'Public Const ID_REPORTCONTROL_PREVIEWMODE = 165
'Public Const ID_REPORTCONTROL_GROUPBYBOX = 166
'Public Const ID_REPORTCONTROL_SHADEGROUPHEADINGS = 167
'Public Const ID_REPORTCONTROL_FIELDCHOOSER = 168
'Public Const ID_REPORTCONTROL_COLLAPSE_ALL_ROWS = 175
'Public Const ID_REPORTCONTROL_EXPAND_ALL_ROWS = 176
'Public Const ID_REPORTCONTROL_AUTOMATIC_FORMAT = 177
'Public Const ID_REPORTCONTROL_GROUP_BY_COLUMN = 178
'Public Const ID_REPORTCONTROL_COLUMN_HIDE = 180
'Public Const ID_REPORTCONTROL_COLUMN_ARRANGE_BY = 181
'Public Const ID_REPORTCONTROL_COLUMN_REMOVE = 182
'Public Const ID_REPORTCONTROL_ASCENDING = 183
'Public Const ID_REPORTCONTROL_DESCENDING = 184
'Public Const ID_REPORTCONTROL_ALIGNMENT = 185
'Public Const ID_REPORTCONTROL_FLAT_COLUMN_STYLE = 223
'Public Const ID_REPORTCONTROL_FILTER_TEXT = 251
'
Public Const IDS_ARRANGE_BY = 220

Public Const ID_THEME_CLIENTPANE = 190

Public Const ID_THEME_OFFICE2000_PLAIN = 191
Public Const ID_THEME_OFFICEXP_PLAIN = 192
Public Const ID_THEME_OFFICE2003_PLAIN = 193
Public Const ID_THEME_NATIVE_PLAIN = 194

'Public Const ID_CALENDAR_NEW = 500
'Public Const ID_CALENDAR_PRINT = 501
'Public Const ID_CALENDAR_DELETE = 502
'Public Const ID_CALENDAR_COLORING = 503
'Public Const ID_CALENDAR_SCHEDULES = 504
'Public Const ID_CALENDAR_TODAY = 505
'Public Const ID_CALENDAR_DAY = 506
'Public Const ID_CALENDAR_WORKWEEK = 507
'Public Const ID_CALENDAR_WEEK = 508
'Public Const ID_CALENDAR_MONTH = 509
'Public Const ID_CALENDAR_FIND = 510
'Public Const ID_CALENDAR_ADDRESSBOOK = 511
'Public Const ID_CALENDAR_HELP = 512
'
Public Const ID_CALENDAREVENT_OPEN = 6050
Public Const ID_CALENDAREVENT_DELETE = 6051
Public Const ID_CALENDAREVENT_NEW = 6052
Public Const ID_CALENDAREVENT_CHANGE_TIMEZONE = 6053
Public Const ID_CALENDAREVENT_60 = 6054
Public Const ID_CALENDAREVENT_30 = 6055
Public Const ID_CALENDAREVENT_15 = 6056
Public Const ID_CALENDAREVENT_10 = 6057
Public Const ID_CALENDAREVENT_5 = 6058

Public Const ID_INDICATOR_CAPS = 59137
Public Const ID_INDICATOR_NUM = 59138
Public Const ID_INDICATOR_SCRL = 59139

Public Const FCONTROL = 8

'Report Control public constants

'public constants used to identify columns, this will be the column ItemIndex
Public Const COLUMN_IMPORTANCE = 0
Public Const COLUMN_ICON = 1
Public Const COLUMN_ATTACHMENT = 2
Public Const COLUMN_FROM = 3
Public Const COLUMN_SUBJECT = 4
Public Const COLUMN_SENT = 5
Public Const COLUMN_SIZE = 6
Public Const COLUMN_CHECK = 7
Public Const COLUMN_PRICE = 8
Public Const COLUMN_CREATED = 9
Public Const COLUMN_RECEIVED = 10
Public Const COLUMN_CONVERSATION = 11
Public Const COLUMN_CONTACTS = 12
Public Const COLUMN_MESSAGE = 13
Public Const COLUMN_CC = 14
Public Const COLUMN_CATEGORIES = 15
Public Const COLUMN_AUTOFORWARD = 16
Public Const COLUMN_DO_NOT_AUTOARCH = 17
Public Const COLUMN_DUE_BY = 18
  
'public constants used to identify icons used in the ReportControl
Public Const COLUMN_MAIL_ICON = 1
Public Const COLUMN_IMPORTANCE_ICON = 2
Public Const COLUMN_CHECK_ICON = 3
Public Const RECORD_UNREAD_MAIL_ICON = 4
Public Const RECORD_READ_MAIL_ICON = 5
Public Const RECORD_REPLIED_ICON = 6
Public Const RECORD_IMPORTANCE_HIGH_ICON = 7
Public Const COLUMN_ATTACHMENT_ICON = 8
Public Const COLUMN_ATTACHMENT_NORMAL_ICON = 9
Public Const RECORD_IMPORTANCE_LOW_ICON = 10

Public Const IMPORTANCE_HIGH = 0
Public Const IMPORTANCE_NORMAL = 1
Public Const IMPORTANCE_LOW = 2

Public Const CHECKED_TRUE = 1
Public Const CHECKED_FALSE = 0

Public Const READ_TRUE = 1
Public Const READ_FALSE = 0

Public Const ATTACHMENTS_TRUE = 1
Public Const ATTACHMENTS_FALSE = 0

'Docking Pane Constants
Public Const PANE_SHORTCUTBAR = 1
Public Const PANE_REPORT_CONTROL = 2
Public Const PANE_READING_PANE = 3
Public Const PANE_FINDBAR = 4
Public ShowEventInPane As Boolean

'Shortcutbar constants
Public Const SHORTCUT_INBOX = 4300
Public Const SHORTCUT_CALENDAR = 4301
Public Const SHORTCUT_CONTACTS = 4302
Public Const SHORTCUT_TASKS = 4303
Public Const SHORTCUT_NOTES = 4304
Public Const SHORTCUT_FOLDER_LIST = 4305
Public Const SHORTCUT_SHORTCUTS = 4306
Public Const SHORTCUT_JOURNAL = 4307

Public Const SHORTCUT_SHOW_MORE = 4308
Public Const SHORTCUT_SHOW_FEWER = 4309

Public Const SHORTCUT_NAVIGATE_PANE_OPTIONS = 4310
Public Const SHORTCUT_ADD_REMOVE_BUTTONS = 4311

'Email Page constants
Public Const ID_EMAIL_SEND = 4330
Public Const ID_EMAIL_ADDRESS_BOOK = 4331
Public Const ID_EMAIL_IMPORTANCE_HIGH = 4332
Public Const ID_EMAIL_ATTACHMENT = 4333
Public Const ID_EMAIL_IMPORTANCE_LOW = 4334
Public Const ID_EMAIL_CHECK_NAMES = 4335
Public Const ID_EMAIL_PERMISSION = 4336
Public Const ID_EMAIL_FLAG = 337
Public Const ID_EMAIL_CREATE_RULE = 338
Public Const ID_EMAIL_OPTIONS = 339
Public Const ID_EMAIL_CLOSE = 340
Public Const ID_EMAIL_SAVE = 341
Public Const ID_EMAIL_BOLD = 346
Public Const ID_EMAIL_ITALIC = 347
Public Const ID_EMAIL_UNDERSCORE = 348
Public Const ID_EMAIL_BULLETS = 349
Public Const ID_EMAIL_NUMBERING = 350
Public Const ID_EMAIL_DECREASE_INDENT = 351
Public Const ID_EMAIL_INCREASE_INDENT = 352
Public Const ID_EMAIL_TRANSLATE = 354
Public Const ID_EMAIL_LTR = 355
Public Const ID_EMAIL_RTL = 356
Public Const ID_EMAIL_LEFT = 357
Public Const ID_EMAIL_CENTER = 358
Public Const ID_EMAIL_RIGHT = 359

Public Const ID_EMAIL_FILE_NEW = 370
Public Const ID_EMAIL_FILE_OPEN = 371
Public Const ID_EMAIL_FILE_SAVE_AS = 372
Public Const ID_EMAIL_FILE_PRINT_SETUP = 373
Public Const ID_EMAIL_FILE_PRINT_PREVIEW = 374
Public Const ID_EMAIL_FILE_PRINT = 375
Public Const ID_EMAIL_FILE_EXIT = 376

'Public Const ID_EMAIL_EDIT_UNDO = 400
'Public Const ID_EMAIL_EDIT_REDO = 401
'Public Const ID_EMAIL_EDIT_CUT = 402
'Public Const ID_EMAIL_EDIT_COPY = 403
'Public Const ID_EMAIL_EDIT_OFFICE_CLIPBOARD = 404
'Public Const ID_EMAIL_EDIT_PASTE = 405
'Public Const ID_EMAIL_EDIT_PASTE_SPECIAL = 406
'Public Const ID_EMAIL_EDIT_PASTE_AS_HYPERLINK = 407
'Public Const ID_EMAIL_EDIT_CLEAR = 408
'Public Const ID_EMAIL_EDIT_SELECT_ALL = 409
'Public Const ID_EMAIL_EDIT_FIND = 410
'Public Const ID_EMAIL_EDIT_REPLACE = 411
'Public Const ID_EMAIL_EDIT_GO_TO = 412
'Public Const ID_EMAIL_EDIT_UPDATE_IME_DICTIONARY = 413
'Public Const ID_EMAIL_EDIT_RECONVERT = 414
'Public Const ID_EMAIL_EDIT_LINKS = 415
'Public Const ID_EMAIL_EDIT_OBJECT = 416

Public Const ID_EMAIL_VIEW_NORMAL = 417
Public Const ID_EMAIL_VIEW_WEB_LAYOUT = 418
Public Const ID_EMAIL_VIEW_PRINT_LAYOUT = 419
Public Const ID_EMAIL_VIEW_READING_LAYOUT = 420
Public Const ID_EMAIL_VIEW_OUTLINE = 421
Public Const ID_EMAIL_VIEW_TASK_PANE = 422
Public Const ID_EMAIL_VIEW_TOOLBARS = 423
Public Const ID_EMAIL_VIEW_RULER = 424
Public Const ID_EMAIL_VIEW_SHOW_PARAGRAPHMARKS = 425
Public Const ID_EMAIL_VIEW_GRIDLINES = 426
Public Const ID_EMAIL_VIEW_DOCUMENT_MAP = 427
Public Const ID_EMAIL_VIEW_THUMBNAILS = 428
Public Const ID_EMAIL_VIEW_HEADER_AND_FOOTER = 429
Public Const ID_EMAIL_VIEW_FOOTNOTES = 430
Public Const ID_EMAIL_VIEW_MARKUP = 431
Public Const ID_EMAIL_VIEW_FULL_SCREEN = 432
Public Const ID_EMAIL_VIEW_ZOOM = 433

Public Const ID_EMAIL_INSERT_BREAK = 434
Public Const ID_EMAIL_INSERT_PAGE_NUMBERS = 435
Public Const ID_EMAIL_INSERT_DATE_AND_TIME = 436
Public Const ID_EMAIL_INSERT_AUTOTEXT = 437
Public Const ID_EMAIL_INSERT_FIELD = 438
Public Const ID_EMAIL_INSERT_SYMBOL = 439
Public Const ID_EMAIL_INSERT_COMMENT = 440
Public Const ID_EMAIL_INSERT_NUMBER = 441
Public Const ID_EMAIL_INSERT_REFERENCE = 442
Public Const ID_EMAIL_INSERT_WEBCOMPONENT = 443
Public Const ID_EMAIL_INSERT_PICTURE = 444
Public Const ID_EMAIL_INSERT_DIAGRAM = 445
Public Const ID_EMAIL_INSERT_TEXT_BOX = 446
Public Const ID_EMAIL_INSERT_FILE = 447
Public Const ID_EMAIL_INSERT_OBJECT = 448
Public Const ID_EMAIL_INSERT_BOOKMARK = 449
Public Const ID_EMAIL_INSERT_HYPERLINK = 450

Public Const ID_EMAIL_FORMAT_FONT = 451
Public Const ID_EMAIL_FORMAT_PARAGRAPH = 452
Public Const ID_EMAIL_FORMAT_BULLETS_AND_NUMBERING = 453
Public Const ID_EMAIL_FORMAT_BORDERS_AND_SHADING = 454
Public Const ID_EMAIL_FORMAT_COLUMNS = 455
Public Const ID_EMAIL_FORMAT_TABS = 456
Public Const ID_EMAIL_FORMAT_DROP_CAP = 457
Public Const ID_EMAIL_FORMAT_TEXT_DIRECTION = 458
Public Const ID_EMAIL_FORMAT_CHANGE_CASE = 459
Public Const ID_EMAIL_FORMAT_FIT_TEXT = 460
Public Const ID_EMAIL_FORMAT_ASIAN_LAYOUT = 461
Public Const ID_EMAIL_FORMAT_BACKGROUND = 462
Public Const ID_EMAIL_FORMAT_THEME = 463
Public Const ID_EMAIL_FORMAT_FRAMES = 464
Public Const ID_EMAIL_FORMAT_AUTOFORMAT = 465
Public Const ID_EMAIL_FORMAT_STYLES_AND_FORMATTING = 466
Public Const ID_EMAIL_FORMAT_REVEAL_FORMATTING = 467
Public Const ID_EMAIL_FORMAT_FORMAT_AUTOSHAPE_PICTURE = 468

Public Const ID_EMAIL_TOOLS_SPELLING_AND_GRAMMAR = 469
Public Const ID_EMAIL_TOOLS_RESEARCH = 470
Public Const ID_EMAIL_TOOLS_LANGUAGE = 471
Public Const ID_EMAIL_TOOLS_FIX_BROKEN_TEXT = 472
Public Const ID_EMAIL_TOOLS_WORDCOUNT = 473
Public Const ID_EMAIL_TOOLS_AUTOSUMMARIZE = 474
Public Const ID_EMAIL_TOOLS_SPEECH = 475
Public Const ID_EMAIL_TOOLS_SHAREDWORKSPACE = 476
Public Const ID_EMAIL_TOOLS_TRACK_CHANGES = 477
Public Const ID_EMAIL_TOOLS_COMPARE_AND_MERGE_DOCUMENTS = 478
Public Const ID_EMAIL_TOOLS_PROTECT_DOCUMENT = 479
Public Const ID_EMAIL_TOOLS_ONLINE_COLLABORATION = 480
Public Const ID_EMAIL_TOOLS_LETTERS_AND_MAILINGS = 481
Public Const ID_EMAIL_TOOLS_MACRO = 482
Public Const ID_EMAIL_TOOLS_TEMPLATES_AND_ADDINS = 483
Public Const ID_EMAIL_TOOLS_AUTOCORRECT_OPTIONS = 484
Public Const ID_EMAIL_TOOLS_CUSTOMIZE = 485
Public Const ID_EMAIL_TOOLS_OPTIONS = 486

Public Const ID_EMAIL_TABLE_DRAW_TABLE = 487
Public Const ID_EMAIL_TABLE_INSERT = 488
Public Const ID_EMAIL_TABLE_DELETE = 489
Public Const ID_EMAIL_TABLE_SELECT = 490
Public Const ID_EMAIL_TABLE_MERGE_CELLS = 491
Public Const ID_EMAIL_TABLE_SPLIT_CELLS = 492
Public Const ID_EMAIL_TABLE_SPLITTABLE = 493
Public Const ID_EMAIL_TABLE_TABLE_AUTOFORMAT = 494
Public Const ID_EMAIL_TABLE_AUTOFIT = 495
Public Const ID_EMAIL_TABLE_HEADING_ROWS_REPEAT = 496
Public Const ID_EMAIL_TABLE_CONVERT = 497
Public Const ID_EMAIL_TABLE_SORT = 498
Public Const ID_EMAIL_TABLE_FORMULA = 499
Public Const ID_EMAIL_TABLE_SHOW_GRIDLINES = 500
Public Const ID_EMAIL_TABLE_TABLE_PROPERTIES = 501

Public Const ID_EMAIL_WINDOW_ARRANGE_ALL = 502
Public Const ID_EMAIL_WINDOW_COMPARE_SIDE_BY_SIDE_WITH = 503
Public Const ID_EMAIL_WINDOW_SPLIT = 504

Public Const ID_EMAIL_HELP_MICROSOFT_OFFICE_WORD_HELP = 505
Public Const ID_EMAIL_HELP_SHOW_THE_OFFICE_ASSISTANT = 506
Public Const ID_EMAIL_HELP_MICROSOFT_OFFICE_ONLINE = 507
Public Const ID_EMAIL_HELP_CONTACT_US = 508
Public Const ID_EMAIL_HELP_WORDPERFECT_HELP = 509
Public Const ID_EMAIL_HELP_CHECK_FOR_UPDATES = 510
Public Const ID_EMAIL_HELP_DETECT_AND_REPAIR = 511
Public Const ID_EMAIL_HELP_ACTIVATE_PRODUCT = 512
Public Const ID_EMAIL_HELP_CUSTOMER_FEEDBACK_OPTIONS = 513
Public Const ID_EMAIL_HELP_ABOUT_MICROSOFT_OFFICE_WORD = 514

Public Const ID_FINDBAR_COMBO = 600
Public Const ID_FINDBAR_SEARCHIN = 601
Public Const ID_FINDBAR_EDIT = 602
Public Const ID_FINDBAR_FINDNOW = 603
Public Const ID_FINDBAR_CLEAR = 604
Public Const ID_FINDBAR_OPTIONS = 605
Public Const ID_FINDBAR_CLOSE = 606






Public Const ID_TAB_HOME = 130

Public Const ID_TAB_EDIT = 133
Public Const ID_TAB_PRINT_PREVIEW = 134


Public Const ID_GROUP_FILE = 130
Public Const ID_GROUP_DOCUMENTVIEWS = 134

Public Const ID_GROUP_WINDOW = 136


Public Const ID_GROUP_EDITING = 139

Public Const ID_VIEW_NORMAL = 141
Public Const ID_VIEW_FULLSCREEN = 142
Public Const ID_WINDOW_SWITCH = 143



Public Const ID_VIEW_WORKSPACE = 59394
Public Const ID_PREVIEW_PRINT_PRINT = 5050
Public Const ID_PREVIEW_PRINT_OPTIONS = 5051
Public Const ID_PREVIEW_PAGESETUP_MARGINS = 5052
Public Const ID_PREVIEW_PAGESETUP_ORIENTATION = 5053
Public Const ID_PREVIEW_PAGESETUP_SIZE = 5054
Public Const ID_PREVIEW_ZOOM_ZOOM = 5055
Public Const ID_PREVIEW_ZOOM_100_PERCENT = 5056
Public Const ID_PREVIEW_ZOOM_1PAGE = 5057
Public Const ID_PREVIEW_ZOOM_2PAGES = 5058
Public Const ID_PREVIEW_ZOOM_PAGE_WIDTH = 5059
Public Const ID_PREVIEW_PREVIEW_RULER = 5060
Public Const ID_PREVIEW_PREVIEW_MAGNIFIER = 5061
Public Const ID_PREVIEW_PREVIEW_SHRINK = 5062
Public Const ID_PREVIEW_PREVIEW_NEXT = 5063
Public Const ID_PREVIEW_PREVIEW_PREVIOUS = 5064
Public Const ID_PREVIEW_PREVIEW_CLOSE = 5065
Public Const ID_GROUP_PREVIEW = 5070
Public Const ID_GROUP_ZOOM = 5071
Public Const ID_GROUP_PRINT = 5072
Public Const ID_MARGINS_CUSTOM_MARGINS = 5073
Public Const ID_ORIENTATION_PORTRAIT = 5074
Public Const ID_ORIENTATION_LANDSCAPE = 5075
Public Const ID_SIZE_MORE_PAPER_SIZES = 5076

Public Const ID_VIEW_MESSAGEBAR = 2815
Public Const ID_GROUP_ADVANCED = 3431
Public Const ID_GROUP_HYPERLINK = 3432
Public Const ID_GROUP_MARKUPLABEL = 3433
Public Const ID_GROUP_BITMAP = 3434
Public Const ID_TAB_ADVANCED = 3435

Public Const ID_VIEW_STATUS_BAR = 2808

Public Const ID_TAB_CALENDAR_HOME = 12000
Public Const ID_GROUP_NEW = 12001
Public Const ID_GROUP_NEW_APPOINTMENT = 12002
Public Const ID_GROUP_NEW_MEETING = 12003
Public Const ID_GROUP_NEW_ITEMS = 12044
Public Const ID_GROUP_NEW_ALLDAY = 12051
Public Const ID_GROUP_GOTO = 12005
Public Const ID_GROUP_GOTO_TODAY = 12006
Public Const ID_GROUP_GOTO_NEXT7DAYS = 12007
Public Const ID_GROUP_ARRANGE2 = 12008
Public Const ID_GROUP_ARRANGE_DAY = 12009
Public Const ID_GROUP_ARRANGE_WORK_WEEK = 12010
Public Const ID_GROUP_ARRANGE_WEEK = 12012
Public Const ID_GROUP_ARRANGE_MONTH = 12012
Public Const ID_GROUP_ARRANGE_MONTH_LOW = 12052
Public Const ID_GROUP_ARRANGE_MONTH_MEDIUM = 12053
Public Const ID_GROUP_ARRANGE_MONTH_HIGH = 12054
Public Const ID_GROUP_ARRANGE_SCHEDULE_VIEW = 12013
Public Const ID_GROUP_MANAGE = 12023
Public Const ID_GROUP_MANAGE_CALENDARS_OPEN = 12014
Public Const ID_GROUP_MANAGE_CALENDARS_GROUPS = 12015
Public Const ID_GROUP_SHARE = 12024
Public Const ID_GROUP_SHARE_EMAIL = 12016
Public Const ID_GROUP_SHARE_SHARE = 12017
Public Const ID_GROUP_SHARE_PUBLISH = 12018
Public Const ID_GROUP_SHARE_PERMISSIONS = 12019
Public Const ID_GROUP_FIND2 = 12020
Public Const ID_GROUP_FIND2_CONTACT = 12021
Public Const ID_GROUP_FIND2_ADDRESSBOOK = 12022

Public Const ID_TAB_MAIL_HOME = 12114
Public Const ID_TAB_SEND_RECEIVE = 12110
Public Const ID_TAB_FOLDER = 12111
Public Const ID_TAB_VIEW2 = 12112
Public Const ID_TAB_ADDINS2 = 12113


Public Const FSHIFT = 4
Public Const FALT = 16

Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B



Public Const ID_WINDOW_NEW = 57648
Public Const ID_WINDOW_ARRANGE = 57649






'****************************************************************************************
'****************************************************************************************
'****************************************************************************************
'
'
'                   Constantes de ARITAXI6
'
'
'****************************************************************************************
'****************************************************************************************
 
Public Const ID_Empresa = 101
Public Const ID_ParametrosContabilidad = 102
Public Const ID_Contadores = 103
Public Const ID_Informes = 104
Public Const ID_Usuarios = 105
Public Const ID_SelImpresora = 106
Public Const ID_ConfigurarBalances = 108

'
Public Const ID_Marcas = 201
Public Const ID_AlmacenesPropios = 202
Public Const ID_TiposUnidad = 203
Public Const ID_TiposArticulos = 204
Public Const ID_FamiliasArticulos = 205
Public Const ID_Articulos = 206
Public Const ID_MovimientosAlmacen = 207
Public Const ID_HistoricoMovimientosAlmacen = 208
Public Const ID_MovimientosArticulos = 209
Public Const ID_ListadoMovimientos = 211
Public Const ID_ArticulosInactivos = 212
Public Const ID_ArticulosComponentes = 213
Public Const ID_ValoracionStocks = 214
Public Const ID_StocksMaxMin = 215
Public Const ID_StocksFecha = 216
Public Const ID_TomaInventario = 217
Public Const ID_EntradaExisReal = 218
Public Const ID_ListadoDiferencias = 219
Public Const ID_ActDiferencias = 220
Public Const ID_ValStocksInven = 221
Public Const ID_HcoInventario = 222



Public Const ID_Actividades = 301
Public Const ID_ClientesAgrup = 302
Public Const ID_FormasPago = 303
Public Const ID_BancosPropio = 304
Public Const ID_SituEspe = 305
Public Const ID_AgentesCom = 306
Public Const ID_Clientes = 307
Public Const ID_TiposCartas = 308
Public Const ID_Incidencias = 309
Public Const ID_Tarjetas = 310
Public Const ID_ClientesInactivos = 311
Public Const ID_InfClientes = 312
Public Const ID_AltasClientes = 313
Public Const ID_EtiquetasClientes = 314
Public Const ID_CartasClientes = 315
Public Const ID_TraspasoTaxitronic = 316
Public Const ID_HcoLlamadas = 317
Public Const ID_ServiciosAbonados = 318
Public Const ID_FacturacionClientes = 319
Public Const ID_FactuVarClientes = 320
Public Const ID_HcoFacturas = 321
Public Const ID_ReimprimirFras = 322
Public Const ID_ContaFacturas = 323
Public Const ID_FrasRectificativas = 324
Public Const ID_VentasporCliente = 325
Public Const ID_DetalleFacturacion = 326



Public Const ID_Trabajadores = 401
Public Const ID_Vehiculos = 402
Public Const ID_Choferes = 403
Public Const ID_Socios = 404
Public Const ID_HistoricoUves = 405
Public Const ID_EtiquetasSocios = 406
Public Const ID_CartasSocios = 408
Public Const ID_Albaranes = 409
Public Const ID_AlbxArt = 410
Public Const ID_AlbAnulados = 411
Public Const ID_PrevFacturacion = 412
Public Const ID_FacturacionAlb = 413
Public Const ID_FacturasRect = 414
Public Const ID_HcoAlbFra = 415
Public Const ID_ReimprirFras = 416
Public Const ID_ContabFras = 417
Public Const ID_ServSocios = 418
Public Const ID_Liquidacion = 419
Public Const ID_HcoFras = 420
Public Const ID_ReimprFras = 421
Public Const ID_ContabilFras = 422
Public Const ID_RetenSocios = 423
Public Const ID_VtasSocios = 424
Public Const ID_VtasMeses = 425
Public Const ID_VtasFamArt = 426
Public Const ID_DetalleFra = 427


Public Const ID_Proveedores = 501
Public Const ID_ProvVarios = 502
Public Const ID_Direcciones = 503
Public Const ID_InfProveedores = 504
Public Const ID_EtiProveedores = 505
Public Const ID_CartasProv = 506
Public Const ID_PreciosProv = 507
Public Const ID_DtosProv = 508
Public Const ID_PedidosProv = 509
Public Const ID_PedidosAnulados = 510
Public Const ID_MatPdteRecibir = 511
Public Const ID_AlbProveedor = 512
Public Const ID_AlbAnuladosPro = 513
Public Const ID_InfPdteFacturar = 514
Public Const ID_RecepFacturas = 515
Public Const ID_HcoAlbxFra = 516
Public Const ID_ContabFacturas = 517
Public Const ID_ComprasProveedor = 518
Public Const ID_ComprasFamxArt = 519
Public Const ID_InfAlbxProv = 520

' PUBLICIDAD
Public Const ID_FacturarClientes = 601
Public Const ID_FrasRectifCli = 602
Public Const ID_HcoFrasClientes = 603
Public Const ID_FacturacionSocios = 604
Public Const ID_FrasRectifSocios = 605
Public Const ID_HcoFacturasSocios = 606
Public Const ID_ReimprFrasSocios = 607
Public Const ID_ContabFrasSocios = 608

' CUOTAS
Public Const ID_GrarFrasCuotas = 701
Public Const ID_ReimprFrasCuotas = 702
Public Const ID_HcoFrasCuotas = 703
Public Const ID_ContabFrasCuotas = 704
Public Const ID_MtoAlbaranes = 705
Public Const ID_PrevFacturacCuotas = 706
Public Const ID_Facturacion = 707
Public Const ID_FrasRectific = 708

' REPARACIONES
Public Const ID_Reparaciones = 801
Public Const ID_ControlRep = 802
Public Const ID_NrosSerie = 803
Public Const ID_MotivosBajaEquipos = 804
Public Const ID_MotivosPdteRepara = 805
Public Const ID_ServAsistenciaTecnica = 806
Public Const ID_TiposAveria = 807
Public Const ID_TrabajosRealizados = 809
Public Const ID_InfReparacionesDia = 810





Public Const ID_InformeporNIF = 901
Public Const ID_Informeporcuenta = 902
Public Const ID_SituaciónTesoreria = 903


Public Const ID_Traspasodecuentasenapuntes = 1408
Public Const ID_Renumerarregistrosproveedor = 1409
Public Const ID_Aumentardígitoscontables = 1410
Public Const ID_TraspasocodigosdeIVA = 1411
Public Const ID_Accionesrealizadas = 1412
Public Const ID_ImportarFacturasCliente = 1413


Public Const ID_Licencia_Usuario_Final_txt = 2001
Public Const ID_Licencia_Usuario_Final_web = 2002
Public Const ID_Ver_Version_operativa_web = 2003


