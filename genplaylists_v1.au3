#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=F:\GenPlaylists.Exe
#AutoIt3Wrapper_Res_Comment=Rekursives erzeugen von Playlisten ab einem Startverzeichnis
#AutoIt3Wrapper_Res_Description=Erzeugen von Playlisten
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Ralf-Andre Lettau
#AutoIt3Wrapper_Res_Language=1031
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -v 12
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#include <FileConstants.au3>
#include <File.au3>
#include <Array.au3>

OnAutoItExitRegister("term")

;~ Global Const $sSplashTitle = "Playlisten erzeugen"

;~ Global Const $csFldrCreErr = "Ordnerliste konnte nicht erstellt werden. Error_:%d Extended_:%s\n"
;~ Global Const $csFldrSearchMsg = "Ordner suchen"
;~ Global Const $csNoFil = "Keine verwertbaren Dateien in %s gefunden.n"
;~ Global Const $csFldrSortErr = "Ordnerliste konnte nicht sortiert werden. Error_:%s Extended_:%s\n"
;~ Global Const $csFlistCreErr = "Dateiliste %s konnte nicht erstellt werden. Error_:%s Extended_:%s\n"
;~ Global Const $csFlistSortErr = "Dateiliste %s konnte nicht sortiert werden. Error_:%s Extended_:%s\n"
;~ Global Const $csFmoveMsg = "umbennen'%s' to '%s'\n"


Global Const $sSplashTitle = "Create Playlists"
Global Const $csFldrSearchMsg = "Set start folder"
Global Const $csNoFil = "No usable files found in %s\n"

Global Const $csFldrCreErr = "Folderlist could not be created. Error_:%d Extended_:%s\n"
Global Const $csFldrSortErr = "Folderlist could not be sorted. Error_:%s Extended_:%s\n"
Global Const $csFlistCreErr = "Filelist %s could not be created Error_:%s Extended_:%s\n"
Global Const $csFlistSortErr = "Filelist %s could not be sorted. Error_:%s Extended_:%s\n"
Global Const $csFmoveMsg = "rename '%s' to '%s'\n"


genPlaylist()

Func term()
	SplashOff()
EndFunc   ;==>term

Func genPlaylist($sSrcPath = "")


	Local Const $Contributing_artists = 13
	Local Const $iAscending = 0

	Local $aFilelist, $iFilelistUpb
	Local $aFldrlist, $aPathSplit
	Local $sPlaylistName, $sArtist
	Local $fh


	Local $sErr, $sExt

	While $sSrcPath = ""
		$sSrcPath = FileSelectFolder($csFldrSearchMsg, "")
		If @error = 1 Then Return (False)
	WEnd

	$aFldrlist = _FileListToArrayRec($sSrcPath, "*", $FLTAR_FOLDERS, $FLTAR_RECUR, Default, $FLTAR_FULLpath)

	SplashTextOn($sSplashTitle, "", 600, 100, -1, -1, BitOR(4, 16), "", 10)

		$sErr = @error
		$sExt = @extended

	Switch $sErr
		Case 0
			;
		Case 9
			$aFldrlist[0] = "1"
			$aFldrlist[1] = "."
		Case Else
			Msg(sf($csFldrCreErr, $sErr, $sExt))
			Exit 1
	EndSwitch


	For $i = 1 To $aFldrlist[0]
		;	$sFldr=stringright($aFldrlist[$i],StringInStr($aFldrlist[$i],"\",0,2,1))
		;
		; Wieso wird teilweise ein Backslash angehangen, teilweise nicht?
		; bis das geklärt ist, halt dieser Workaround
		;
		$sPlaylistName = StringReplace($aFldrlist[$i] & ".m3u", "\.m3u", ".m3u")

		Msg(sf("%s\n", $sPlaylistName))
		$fh = FileOpen($sPlaylistName, $FO_OVERWRITE)


		$aFilelist = _FileListToArrayRec($aFldrlist[$i], "*", $FLTAR_FILES, $FLTAR_RECUR, Default, $FLTAR_FULLpath)
		$sErr = @error
		$sExt = @extended
		If $sErr Then
			If $sExt = 9 Then
				Msg(sf($csNoFil, $aFldrlist[$i]))
			Else
				Msg(sf($csFlistCreErr, $aFldrlist[$i], $sErr, $sExt))
			EndIf
		Else
			$iFilelistUpb = $aFilelist[0]
			_ArraySort($aFilelist, $iAscending, 1, $iFilelistUpb)
			$sErr = @error
			$sExt = @extended
			If $sErr Then Msg(sf($csFlistSortErr, $sPlaylistName, $sErr, $sExt))
			;
			;  my Toyota Touch needs the #EXTM3U entry
			;  for position independend playlist.
			;
			FileWriteLine($fh, "#EXTM3U")

			For $j = 1 To $aFilelist[0]
				If StringRight($aFilelist[$j], 4) = ".mp3" Then
					FileWriteLine($fh, StringTrimLeft($aFilelist[$j], 2))
				EndIf
			Next
		EndIf
		FileClose($fh)

		;
		; set Artist of the first file in list as Prefix
		;
		If IsArray($aFilelist) Then
			$sArtist = _GetExtProperty($aFilelist[1], $Contributing_artists)
			If @error Or $sArtist = "0" Then
			Else
				$aPathSplit = PathSplit($sPlaylistName)
				;
				; ignore if the atrist is already part of the name
				;
				If StringInStr($aPathSplit[3], $sArtist) > 0 Or (StringLen($aPathSplit[3]) < 3) Then
				Else
					Msg(sf($csFmoveMsg, $sPlaylistName, $aPathSplit[1] & $aPathSplit[2] & $sArtist & "-" & $aPathSplit[3] & $aPathSplit[4]))
					FileMove($sPlaylistName, $aPathSplit[1] & $aPathSplit[2] & $sArtist & "-" & $aPathSplit[3] & $aPathSplit[4])
				EndIf
			EndIf
		EndIf
		;
		;

	Next

EndFunc   ;==>genPlaylist


Func Msg($sMsg)
	ConsoleWrite($sMsg)
	ControlSetText($sSplashTitle, "", "Static1", $sMsg)
EndFunc   ;==>Msg

;===============================================================================
;
; Function Name:   sf($sFormat,$var1...$var20)
; Description:    liefert einen formatierten String zurück
; Parameter(s):    $sFormat - Formatstring
;                  $var1...$var20 -
; Requirement(s):
; Return Value(s):  formatierter String
; Author(s):
; Modified:
; Remarks:			Doku siehe StringFormat

;===============================================================================
Func sf($sFormat, $v1 = "", $v2 = "", $v3 = "", $v4 = "", $v5 = "", $v6 = "", $v7 = "", $v8 = "", $v9 = "", $v10 = "", _
		$v11 = "", $v12 = "", $v13 = "", $v14 = "", $v15 = "", $v66 = "", $v17 = "", $v18 = "", $v19 = "", $v20 = "")
	Return StringFormat($sFormat, $v1, $v2, $v3, $v4, $v5, $v6, $v7, $v8, $v9, $v10, $v11, $v12, $v13, $v14, $v15, $v66, $v17, $v18, $v19, $v20)
EndFunc   ;==>sf


Func PathSplit($sFilePath)

	Local $sDir

	Local $aArray = StringRegExp($sFilePath, "^\h*((?:\\\\\?\\)*(\\\\[^\?\/\\]+|[A-Za-z]:)?(.*[\/\\]\h*)?((?:[^\.\/\\]|(?(?=\.[^\/\\]*\.)\.))*)?([^\/\\]*))$", $STR_REGEXPARRAYMATCH)
	If @error Then ; Just in case.
		ReDim $aArray[5]
		$aArray[0] = $sFilePath
	EndIf

	If StringLeft($aArray[2], 1) == "/" Then
		$sDir = StringRegExpReplace($aArray[2], "\h*[\/\\]+\h*", "\/")
	Else
		$sDir = StringRegExpReplace($aArray[2], "\h*[\/\\]+\h*", "\\")
	EndIf
	$aArray[2] = $sDir

	Return $aArray
EndFunc   ;==>PathSplit

; Function Name:	GetExtProperty($sPath,$iProp)
; Description: Returns an extended property of a given file.
; Parameter(s): $sPath - The path to the file you are attempting to retrieve an extended property from.
; $iProp - The numerical value for the property you want returned. If $iProp is is set
;	to -1 then all properties will be returned in a 1 dimensional array in their corresponding order.
;	The properties are as follows:

; Requirement(s): File specified in $spath must exist.
; Return Value(s): On Success - The extended file property, or if $iProp = -1 then an array with all properties
; On Failure - 0, @Error - 1 (If file does not exist)

; Note(s):
;
;===============================================================================
Func _GetExtProperty($sPath, $iProp)
	Local $iExist, $sFile, $sDir, $oShellApp, $oDir, $oFile, $aProperty, $sProperty
	$iExist = FileExists($sPath)
	If $iExist = 0 Then
		SetError(1)
		Return 0
	Else
		$sFile = StringTrimLeft($sPath, StringInStr($sPath, "\", 0, -1))
		$sDir = StringTrimRight($sPath, (StringLen($sPath) - StringInStr($sPath, "\", 0, -1)))
		$oShellApp = ObjCreate("shell.application")
		$oDir = $oShellApp.NameSpace($sDir)
		$oFile = $oDir.Parsename($sFile)

		If $iProp = -1 Then
			Local $aProperty[35]
			For $i = 0 To 34
				$aProperty[$i] = $oDir.GetDetailsOf($oFile, $i)
			Next
			Return $aProperty
		Else
			$sProperty = $oDir.GetDetailsOf($oFile, $iProp)
			If $sProperty = "" Then
				Return 0
			Else
				Return $sProperty
			EndIf
		EndIf
	EndIf
EndFunc   ;==>_GetExtProperty





#comments-start
	Global Const  $Name=0
	Global Const  $Size=1
	Global Const  $Item_type=2
	Global Const  $Date_modified=3
	Global Const  $Date_created=4
	Global Const  $Date_accessed=5
	Global Const  $Attributes=6
	Global Const  $Offline_status=7
	Global Const  $Offline_availability=8
	Global Const  $Perceived_type=9
	Global Const  $Owner=10
	Global Const  $Kind=11
	Global Const  $Date_taken=12
	Global Const  $Contributing_artists=13
	Global Const  $Album=14
	Global Const  $Year=15
	Global Const  $Genre=16
	Global Const  $Conductors=17
	Global Const  $Tags=18
	Global Const  $Rating=19
	Global Const  $Authors=20
	Global Const  $Title=21
	Global Const  $Subject=22
	Global Const  $Categories=23
	Global Const  $Comments=24
	Global Const  $Copyright=25
	Global Const  $unklar=26
	Global Const  $Length=27
	Global Const  $Bit_rate=28
	Global Const  $Protected=29
	Global Const  $Camera_model=30
	Global Const  $Dimensions=31
	Global Const  $Camera_maker=32
	Global Const  $Company=33
	Global Const  $File_description=34
	Global Const  $Program_name=35
	Global Const  $Duration=36
	Global Const  $Is_online=37
	Global Const  $Is_recurring=38
	Global Const  $Location=39
	Global Const  $Optional_attendee_addresses=40
	Global Const  $Optional_attendees=41
	Global Const  $Organizer_address=42
	Global Const  $Organizer_name=43
	Global Const  $Reminder_time=44
	Global Const  $Required_attendee_addresses=45
	Global Const  $Required_attendees=46
	Global Const  $Resources=47
	Global Const  $Meeting_status=48
	Global Const  $Free_busy_status=49
	Global Const  $Total_size=50
	Global Const  $Account_name=51
	Global Const  $Task_status=52
	Global Const  $Computer=53
	Global Const  $Anniversary=54
	Global Const  $Assistants_name=55
	Global Const  $Assistants_phone=56
	Global Const  $Birthday=57
	Global Const  $Business_address=58
	Global Const  $Business_city=59
	Global Const  $Business_PO_box=60
	Global Const  $Business_postal_code=61
	Global Const  $Business_state_or_province=62
	Global Const  $Business_street=63
	Global Const  $Business_fax=64
	Global Const  $Business_home_page=65
	Global Const  $Business_phone=66
	Global Const  $Callback_number=67
	Global Const  $Car_phone=68
	Global Const  $Children=69
	Global Const  $Company_main_phone=70
	Global Const  $Department=71
	Global Const  $E_mail_address=72
	Global Const  $E_mail2=73
	Global Const  $E_mail3=74
	Global Const  $E_mail_list=75
	Global Const  $E_mail_display_name=76
	Global Const  $File_as=77
	Global Const  $First_name=78
	Global Const  $Full_name=79
	Global Const  $Gender=80
	Global Const  $Given_name=81
	Global Const  $Hobbies=82
	Global Const  $Home_address=83
	Global Const  $Home_city=84
	Global Const  $Home_country_region=85
	Global Const  $Home_PO_box=86
	Global Const  $Home_postal_code=87
	Global Const  $Home_state_or_province=88
	Global Const  $Home_street=89
	Global Const  $Home_fax=90
	Global Const  $Home_phone=91
	Global Const  $IM_addresses=92
	Global Const  $Initials=93
	Global Const  $Job_title=94
	Global Const  $Label=95
	Global Const  $Last_name=96
	Global Const  $Mailing_address=97
	Global Const  $Middle_name=98
	Global Const  $Cell_phone=99
	Global Const  $Cell_phone1=100
	Global Const  $Nickname=101
	Global Const  $Office_location=102
	Global Const  $Other_address=103
	Global Const  $Other_city=104
	Global Const  $Other_country_region=105
	Global Const  $Other_PO_box=106
	Global Const  $Other_postal_code=107
	Global Const  $Other_state_or_province=108
	Global Const  $Other_street=109
	Global Const  $Pager=110
	Global Const  $Personal_title=111
	Global Const  $City=112
	Global Const  $Country_region=113
	Global Const  $PO_box=114
	Global Const  $Postal_code=115
	Global Const  $State_or_province=116
	Global Const  $Street=117
	Global Const  $Primary_e_mail=118
	Global Const  $Primary_phone=119
	Global Const  $Profession=120
	Global Const  $Spouse_Partner=121
	Global Const  $Suffix=122
	Global Const  $TTY_TTD_phone=123
	Global Const  $Telex=124
	Global Const  $Webpage=125
	Global Const  $Content_status=126
	Global Const  $Content_type=127
	Global Const  $Date_acquired=128
	Global Const  $Date_archived=129
	Global Const  $Date_completed=130
	Global Const  $Device_category=131
	Global Const  $Connected=132
	Global Const  $Discovery_method=133
	Global Const  $Friendly_name=134
	Global Const  $Local_computer=135
	Global Const  $Manufacturer=136
	Global Const  $Model=137
	Global Const  $Paired=138
	Global Const  $Classification=139
	Global Const  $Status=140
	Global Const  $Client_ID=141
	Global Const  $Contributors=142
	Global Const  $Content_created=143
	Global Const  $Last_ed=144
	Global Const  $Date_last_saved=145
	Global Const  $Division=146
	Global Const  $Document_ID=147
	Global Const  $Pages=148
	Global Const  $Slides=149
	Global Const  $Total_editing_time=150
	Global Const  $Word_count=151
	Global Const  $Due_date=152
	Global Const  $End_date=153
	Global Const  $File_count=154
	Global Const  $Filename=155
	Global Const  $File_version=156
	Global Const  $Flag_color=157
	Global Const  $Flag_status=158
	Global Const  $Space_free=159
	Global Const  $Bit_depth=160
	Global Const  $Horizontal_resolution=161
	Global Const  $Width=162
	Global Const  $Vertical_resolution=163
	Global Const  $Height=164
	Global Const  $Importance=165
	Global Const  $Is_attachment=166
	Global Const  $Is_deleted=167
	Global Const  $Encryption_status=168
	Global Const  $Has_flag=169
	Global Const  $Is_completed=170
	Global Const  $Incomplete=171
	Global Const  $Read_status=172
	Global Const  $Shared=173
	Global Const  $Creators=174
	Global Const  $Date=175
	Global Const  $Folder_name=176
	Global Const  $Folder_path=177
	Global Const  $Folder=178
	Global Const  $Participants=179
	Global Const  $Path=180
	Global Const  $By_location=181
	Global Const  $Type=182
	Global Const  $Contact_names=183
	Global Const  $Entry_type=184
	Global Const  $Language=185
	Global Const  $Date_visited=186
	Global Const  $Description=187
	Global Const  $Link_status=188
	Global Const  $Link_target=189
	Global Const  $URL=190
	Global Const  $Media_created=191
	Global Const  $Date_released=192
	Global Const  $Encoded_by=193
	Global Const  $Producers=194
	Global Const  $Publisher=195
	Global Const  $Subtitle=196
	Global Const  $User_web_URL=197
	Global Const  $Writers=198
	Global Const  $Attachments=199
	Global Const  $Bcc_addresses=200
	Global Const  $Bcc=201
	Global Const  $Cc_addresses=202
	Global Const  $Cc=203
	Global Const  $Conversation_ID=204
	Global Const  $Date_received=205
	Global Const  $Date_sent=206
	Global Const  $From_addresses=207
	Global Const  $From=208
	Global Const  $Has_attachments=209
	Global Const  $Sender_address=210
	Global Const  $Sender_name=211
	Global Const  $Store=212
	Global Const  $To_addresses=213
	Global Const  $To_do_title=214
	Global Const  $To=215
	Global Const  $Mileage=216
	Global Const  $Album_artist=217
	Global Const  $Album_ID=218
	Global Const  $Beats_per_minute=219
	Global Const  $Composers=220
	Global Const  $Initial_key=221
	Global Const  $Part_of_a_compilation=222
	Global Const  $Mood=223
	Global Const  $Part_of_set=224
	Global Const  $Period=225
	Global Const  $Color=226
	Global Const  $Parental_rating=227
	Global Const  $Parental_rating_reason=228
	Global Const  $Space_used=229
	Global Const  $EXIF_version=230
	Global Const  $Event=231
	Global Const  $Exposure_bias=232
	Global Const  $Exposure_program=233
	Global Const  $Exposure_time=234
	Global Const  $F_stop=235
	Global Const  $Flash_mode=236
	Global Const  $Focal_length=237
	Global Const  $35mm_focal_length=238
	Global Const  $ISO_speed=239
	Global Const  $Lens_maker=240
	Global Const  $Lens_model=241
	Global Const  $Light_source=242
	Global Const  $Max_aperture=243
	Global Const  $Metering_mode=244
	Global Const  $Orientation=245
	Global Const  $People=246
	Global Const  $Program_mode=247
	Global Const  $Saturation=248
	Global Const  $Subject_distance=249
	Global Const  $White_balance=250
	Global Const  $Priority=251
	Global Const  $Project=252
	Global Const  $Channel_number=253
	Global Const  $Episode_name=254
	Global Const  $Closed_captioning=255
	Global Const  $Rerun=256
	Global Const  $SAP=257
	Global Const  $Broadcast_ate=258
	Global Const  $Program_description=259
	Global Const  $Recording_time=260
	Global Const  $Station_call_ign=261
	Global Const  $Station_name=262
	Global Const  $Summary=263
	Global Const  $Snippets=264
	Global Const  $Auto_summary=265
	Global Const  $Search_ranking=266
	Global Const  $Sensitivity=267
	Global Const  $Shared_with=268
	Global Const  $Sharing_status=269
	Global Const  $Product_name=270
	Global Const  $Product_version=271
	Global Const  $Supportlink=272
	Global Const  $Source=273
	Global Const  $Startdate=274
	Global Const  $Billing_information=275
	Global Const  $Complete=276
	Global Const  $Task_owner=277
	Global Const  $Total_file_size=278
	Global Const  $Legal_trademarks=279
	Global Const  $Video_compression=280
	Global Const  $Directors=281
	Global Const  $Data_rate=282
	Global Const  $Frame_height=283
	Global Const  $Frame_rate=284
	Global Const  $Frame_width=285
	Global Const  $Total_bitrate=286
	Global Const  $Masters_Keywords=287
	Global Const  $Masters_Keywords1=288

#comments-end
