# -*- coding: mbcs -*-
# Created by makepy.py forked at https://github.com/tonyroberts/pywin32/tree/type-stubs
# From type library 'MSO.DLL'
# On Tue Jan 12 16:25:37 2021

from __future__ import annotations

'Microsoft Office 16.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x30700f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from enum import Enum
import typing
from pywintypes import IID
from win32com.client import Dispatch


# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}')
MajorVersion = 2
MinorVersion = 8
LibraryFlags = 8
LCID = 0x0

class BackstageGroupStyle(Enum):
	BackstageGroupStyleNormal     =0         
	BackstageGroupStyleWarning    =1         
	BackstageGroupStyleError      =2         

class CertificateDetail(Enum):
	certdetAvailable              =0         
	certdetSubject                =1         
	certdetIssuer                 =2         
	certdetExpirationDate         =3         
	certdetThumbprint             =4         

class CertificateVerificationResults(Enum):
	certverresError               =0         
	certverresVerifying           =1         
	certverresUnverified          =2         
	certverresValid               =3         
	certverresInvalid             =4         
	certverresExpired             =5         
	certverresRevoked             =6         
	certverresUntrusted           =7         

class ContentVerificationResults(Enum):
	contverresError               =0         
	contverresVerifying           =1         
	contverresUnverified          =2         
	contverresValid               =3         
	contverresModified            =4         

class DocProperties(Enum):
	offPropertyTypeNumber         =1         
	offPropertyTypeBoolean        =2         
	offPropertyTypeDate           =3         
	offPropertyTypeString         =4         
	offPropertyTypeFloat          =5         

class EncryptionCipherMode(Enum):
	cipherModeECB                 =0         
	cipherModeCBC                 =1         

class EncryptionProviderDetail(Enum):
	encprovdetUrl                 =0         
	encprovdetAlgorithm           =1         
	encprovdetBlockCipher         =2         
	encprovdetCipherBlockSize     =3         
	encprovdetCipherMode          =4         

class MailFormat(Enum):
	mfPlainText                   =1         
	mfHTML                        =2         
	mfRTF                         =3         

class MsoAlertButtonType(Enum):
	msoAlertButtonOK              =0         
	msoAlertButtonOKCancel        =1         
	msoAlertButtonAbortRetryIgnore=2         
	msoAlertButtonYesNoCancel     =3         
	msoAlertButtonYesNo           =4         
	msoAlertButtonRetryCancel     =5         
	msoAlertButtonYesAllNoCancel  =6         

class MsoAlertCancelType(Enum):
	msoAlertCancelDefault         =-1        
	msoAlertCancelFirst           =0         
	msoAlertCancelSecond          =1         
	msoAlertCancelThird           =2         
	msoAlertCancelFourth          =3         
	msoAlertCancelFifth           =4         

class MsoAlertDefaultType(Enum):
	msoAlertDefaultFirst          =0         
	msoAlertDefaultSecond         =1         
	msoAlertDefaultThird          =2         
	msoAlertDefaultFourth         =3         
	msoAlertDefaultFifth          =4         

class MsoAlertIconType(Enum):
	msoAlertIconNoIcon            =0         
	msoAlertIconCritical          =1         
	msoAlertIconQuery             =2         
	msoAlertIconWarning           =3         
	msoAlertIconInfo              =4         

class MsoAlignCmd(Enum):
	msoAlignLefts                 =0         
	msoAlignCenters               =1         
	msoAlignRights                =2         
	msoAlignTops                  =3         
	msoAlignMiddles               =4         
	msoAlignBottoms               =5         

class MsoAnimationType(Enum):
	msoAnimationIdle              =1         
	msoAnimationGreeting          =2         
	msoAnimationGoodbye           =3         
	msoAnimationBeginSpeaking     =4         
	msoAnimationRestPose          =5         
	msoAnimationCharacterSuccessMajor=6         
	msoAnimationGetAttentionMajor =11        
	msoAnimationGetAttentionMinor =12        
	msoAnimationSearching         =13        
	msoAnimationPrinting          =18        
	msoAnimationGestureRight      =19        
	msoAnimationWritingNotingSomething=22        
	msoAnimationWorkingAtSomething=23        
	msoAnimationThinking          =24        
	msoAnimationSendingMail       =25        
	msoAnimationListensToComputer =26        
	msoAnimationDisappear         =31        
	msoAnimationAppear            =32        
	msoAnimationGetArtsy          =100       
	msoAnimationGetTechy          =101       
	msoAnimationGetWizardy        =102       
	msoAnimationCheckingSomething =103       
	msoAnimationLookDown          =104       
	msoAnimationLookDownLeft      =105       
	msoAnimationLookDownRight     =106       
	msoAnimationLookLeft          =107       
	msoAnimationLookRight         =108       
	msoAnimationLookUp            =109       
	msoAnimationLookUpLeft        =110       
	msoAnimationLookUpRight       =111       
	msoAnimationSaving            =112       
	msoAnimationGestureDown       =113       
	msoAnimationGestureLeft       =114       
	msoAnimationGestureUp         =115       
	msoAnimationEmptyTrash        =116       

class MsoAppLanguageID(Enum):
	msoLanguageIDInstall          =1         
	msoLanguageIDUI               =2         
	msoLanguageIDHelp             =3         
	msoLanguageIDExeMode          =4         
	msoLanguageIDUIPrevious       =5         

class MsoArrowheadLength(Enum):
	msoArrowheadLengthMixed       =-2        
	msoArrowheadShort             =1         
	msoArrowheadLengthMedium      =2         
	msoArrowheadLong              =3         

class MsoArrowheadStyle(Enum):
	msoArrowheadStyleMixed        =-2        
	msoArrowheadNone              =1         
	msoArrowheadTriangle          =2         
	msoArrowheadOpen              =3         
	msoArrowheadStealth           =4         
	msoArrowheadDiamond           =5         
	msoArrowheadOval              =6         

class MsoArrowheadWidth(Enum):
	msoArrowheadWidthMixed        =-2        
	msoArrowheadNarrow            =1         
	msoArrowheadWidthMedium       =2         
	msoArrowheadWide              =3         

class MsoAssignmentMethod(Enum):
	NOT_SET                       =-1        
	STANDARD                      =0         
	PRIVILEGED                    =1         
	AUTO                          =2         

class MsoAutoShapeType(Enum):
	msoShapeMixed                 =-2        
	msoShapeRectangle             =1         
	msoShapeParallelogram         =2         
	msoShapeTrapezoid             =3         
	msoShapeDiamond               =4         
	msoShapeRoundedRectangle      =5         
	msoShapeOctagon               =6         
	msoShapeIsoscelesTriangle     =7         
	msoShapeRightTriangle         =8         
	msoShapeOval                  =9         
	msoShapeHexagon               =10        
	msoShapeCross                 =11        
	msoShapeRegularPentagon       =12        
	msoShapeCan                   =13        
	msoShapeCube                  =14        
	msoShapeBevel                 =15        
	msoShapeFoldedCorner          =16        
	msoShapeSmileyFace            =17        
	msoShapeDonut                 =18        
	msoShapeNoSymbol              =19        
	msoShapeBlockArc              =20        
	msoShapeHeart                 =21        
	msoShapeLightningBolt         =22        
	msoShapeSun                   =23        
	msoShapeMoon                  =24        
	msoShapeArc                   =25        
	msoShapeDoubleBracket         =26        
	msoShapeDoubleBrace           =27        
	msoShapePlaque                =28        
	msoShapeLeftBracket           =29        
	msoShapeRightBracket          =30        
	msoShapeLeftBrace             =31        
	msoShapeRightBrace            =32        
	msoShapeRightArrow            =33        
	msoShapeLeftArrow             =34        
	msoShapeUpArrow               =35        
	msoShapeDownArrow             =36        
	msoShapeLeftRightArrow        =37        
	msoShapeUpDownArrow           =38        
	msoShapeQuadArrow             =39        
	msoShapeLeftRightUpArrow      =40        
	msoShapeBentArrow             =41        
	msoShapeUTurnArrow            =42        
	msoShapeLeftUpArrow           =43        
	msoShapeBentUpArrow           =44        
	msoShapeCurvedRightArrow      =45        
	msoShapeCurvedLeftArrow       =46        
	msoShapeCurvedUpArrow         =47        
	msoShapeCurvedDownArrow       =48        
	msoShapeStripedRightArrow     =49        
	msoShapeNotchedRightArrow     =50        
	msoShapePentagon              =51        
	msoShapeChevron               =52        
	msoShapeRightArrowCallout     =53        
	msoShapeLeftArrowCallout      =54        
	msoShapeUpArrowCallout        =55        
	msoShapeDownArrowCallout      =56        
	msoShapeLeftRightArrowCallout =57        
	msoShapeUpDownArrowCallout    =58        
	msoShapeQuadArrowCallout      =59        
	msoShapeCircularArrow         =60        
	msoShapeFlowchartProcess      =61        
	msoShapeFlowchartAlternateProcess=62        
	msoShapeFlowchartDecision     =63        
	msoShapeFlowchartData         =64        
	msoShapeFlowchartPredefinedProcess=65        
	msoShapeFlowchartInternalStorage=66        
	msoShapeFlowchartDocument     =67        
	msoShapeFlowchartMultidocument=68        
	msoShapeFlowchartTerminator   =69        
	msoShapeFlowchartPreparation  =70        
	msoShapeFlowchartManualInput  =71        
	msoShapeFlowchartManualOperation=72        
	msoShapeFlowchartConnector    =73        
	msoShapeFlowchartOffpageConnector=74        
	msoShapeFlowchartCard         =75        
	msoShapeFlowchartPunchedTape  =76        
	msoShapeFlowchartSummingJunction=77        
	msoShapeFlowchartOr           =78        
	msoShapeFlowchartCollate      =79        
	msoShapeFlowchartSort         =80        
	msoShapeFlowchartExtract      =81        
	msoShapeFlowchartMerge        =82        
	msoShapeFlowchartStoredData   =83        
	msoShapeFlowchartDelay        =84        
	msoShapeFlowchartSequentialAccessStorage=85        
	msoShapeFlowchartMagneticDisk =86        
	msoShapeFlowchartDirectAccessStorage=87        
	msoShapeFlowchartDisplay      =88        
	msoShapeExplosion1            =89        
	msoShapeExplosion2            =90        
	msoShape4pointStar            =91        
	msoShape5pointStar            =92        
	msoShape8pointStar            =93        
	msoShape16pointStar           =94        
	msoShape24pointStar           =95        
	msoShape32pointStar           =96        
	msoShapeUpRibbon              =97        
	msoShapeDownRibbon            =98        
	msoShapeCurvedUpRibbon        =99        
	msoShapeCurvedDownRibbon      =100       
	msoShapeVerticalScroll        =101       
	msoShapeHorizontalScroll      =102       
	msoShapeWave                  =103       
	msoShapeDoubleWave            =104       
	msoShapeRectangularCallout    =105       
	msoShapeRoundedRectangularCallout=106       
	msoShapeOvalCallout           =107       
	msoShapeCloudCallout          =108       
	msoShapeLineCallout1          =109       
	msoShapeLineCallout2          =110       
	msoShapeLineCallout3          =111       
	msoShapeLineCallout4          =112       
	msoShapeLineCallout1AccentBar =113       
	msoShapeLineCallout2AccentBar =114       
	msoShapeLineCallout3AccentBar =115       
	msoShapeLineCallout4AccentBar =116       
	msoShapeLineCallout1NoBorder  =117       
	msoShapeLineCallout2NoBorder  =118       
	msoShapeLineCallout3NoBorder  =119       
	msoShapeLineCallout4NoBorder  =120       
	msoShapeLineCallout1BorderandAccentBar=121       
	msoShapeLineCallout2BorderandAccentBar=122       
	msoShapeLineCallout3BorderandAccentBar=123       
	msoShapeLineCallout4BorderandAccentBar=124       
	msoShapeActionButtonCustom    =125       
	msoShapeActionButtonHome      =126       
	msoShapeActionButtonHelp      =127       
	msoShapeActionButtonInformation=128       
	msoShapeActionButtonBackorPrevious=129       
	msoShapeActionButtonForwardorNext=130       
	msoShapeActionButtonBeginning =131       
	msoShapeActionButtonEnd       =132       
	msoShapeActionButtonReturn    =133       
	msoShapeActionButtonDocument  =134       
	msoShapeActionButtonSound     =135       
	msoShapeActionButtonMovie     =136       
	msoShapeBalloon               =137       
	msoShapeNotPrimitive          =138       
	msoShapeFlowchartOfflineStorage=139       
	msoShapeLeftRightRibbon       =140       
	msoShapeDiagonalStripe        =141       
	msoShapePie                   =142       
	msoShapeNonIsoscelesTrapezoid =143       
	msoShapeDecagon               =144       
	msoShapeHeptagon              =145       
	msoShapeDodecagon             =146       
	msoShape6pointStar            =147       
	msoShape7pointStar            =148       
	msoShape10pointStar           =149       
	msoShape12pointStar           =150       
	msoShapeRound1Rectangle       =151       
	msoShapeRound2SameRectangle   =152       
	msoShapeRound2DiagRectangle   =153       
	msoShapeSnipRoundRectangle    =154       
	msoShapeSnip1Rectangle        =155       
	msoShapeSnip2SameRectangle    =156       
	msoShapeSnip2DiagRectangle    =157       
	msoShapeFrame                 =158       
	msoShapeHalfFrame             =159       
	msoShapeTear                  =160       
	msoShapeChord                 =161       
	msoShapeCorner                =162       
	msoShapeMathPlus              =163       
	msoShapeMathMinus             =164       
	msoShapeMathMultiply          =165       
	msoShapeMathDivide            =166       
	msoShapeMathEqual             =167       
	msoShapeMathNotEqual          =168       
	msoShapeCornerTabs            =169       
	msoShapeSquareTabs            =170       
	msoShapePlaqueTabs            =171       
	msoShapeGear6                 =172       
	msoShapeGear9                 =173       
	msoShapeFunnel                =174       
	msoShapePieWedge              =175       
	msoShapeLeftCircularArrow     =176       
	msoShapeLeftRightCircularArrow=177       
	msoShapeSwooshArrow           =178       
	msoShapeCloud                 =179       
	msoShapeChartX                =180       
	msoShapeChartStar             =181       
	msoShapeChartPlus             =182       
	msoShapeLineInverse           =183       

class MsoAutoSize(Enum):
	msoAutoSizeMixed              =-2        
	msoAutoSizeNone               =0         
	msoAutoSizeShapeToFitText     =1         
	msoAutoSizeTextToFitShape     =2         

class MsoAutomationSecurity(Enum):
	msoAutomationSecurityLow      =1         
	msoAutomationSecurityByUI     =2         
	msoAutomationSecurityForceDisable=3         

class MsoBackgroundStyleIndex(Enum):
	msoBackgroundStyleMixed       =-2        
	msoBackgroundStyleNotAPreset  =0         
	msoBackgroundStylePreset1     =1         
	msoBackgroundStylePreset2     =2         
	msoBackgroundStylePreset3     =3         
	msoBackgroundStylePreset4     =4         
	msoBackgroundStylePreset5     =5         
	msoBackgroundStylePreset6     =6         
	msoBackgroundStylePreset7     =7         
	msoBackgroundStylePreset8     =8         
	msoBackgroundStylePreset9     =9         
	msoBackgroundStylePreset10    =10        
	msoBackgroundStylePreset11    =11        
	msoBackgroundStylePreset12    =12        

class MsoBalloonButtonType(Enum):
	msoBalloonButtonYesToAll      =-15       
	msoBalloonButtonOptions       =-14       
	msoBalloonButtonTips          =-13       
	msoBalloonButtonClose         =-12       
	msoBalloonButtonSnooze        =-11       
	msoBalloonButtonSearch        =-10       
	msoBalloonButtonIgnore        =-9        
	msoBalloonButtonAbort         =-8        
	msoBalloonButtonRetry         =-7        
	msoBalloonButtonNext          =-6        
	msoBalloonButtonBack          =-5        
	msoBalloonButtonNo            =-4        
	msoBalloonButtonYes           =-3        
	msoBalloonButtonCancel        =-2        
	msoBalloonButtonOK            =-1        
	msoBalloonButtonNull          =0         

class MsoBalloonErrorType(Enum):
	msoBalloonErrorNone           =0         
	msoBalloonErrorOther          =1         
	msoBalloonErrorTooBig         =2         
	msoBalloonErrorOutOfMemory    =3         
	msoBalloonErrorBadPictureRef  =4         
	msoBalloonErrorBadReference   =5         
	msoBalloonErrorButtonlessModal=6         
	msoBalloonErrorButtonModeless =7         
	msoBalloonErrorBadCharacter   =8         
	msoBalloonErrorCOMFailure     =9         
	msoBalloonErrorCharNotTopmostForModal=10        
	msoBalloonErrorTooManyControls=11        

class MsoBalloonType(Enum):
	msoBalloonTypeButtons         =0         
	msoBalloonTypeBullets         =1         
	msoBalloonTypeNumbers         =2         

class MsoBarPosition(Enum):
	msoBarLeft                    =0         
	msoBarTop                     =1         
	msoBarRight                   =2         
	msoBarBottom                  =3         
	msoBarFloating                =4         
	msoBarPopup                   =5         
	msoBarMenuBar                 =6         

class MsoBarProtection(Enum):
	msoBarNoProtection            =0         
	msoBarNoCustomize             =1         
	msoBarNoResize                =2         
	msoBarNoMove                  =4         
	msoBarNoChangeVisible         =8         
	msoBarNoChangeDock            =16        
	msoBarNoVerticalDock          =32        
	msoBarNoHorizontalDock        =64        

class MsoBarRow(Enum):
	msoBarRowFirst                =0         
	msoBarRowLast                 =-1        

class MsoBarType(Enum):
	msoBarTypeNormal              =0         
	msoBarTypeMenuBar             =1         
	msoBarTypePopup               =2         

class MsoBaselineAlignment(Enum):
	msoBaselineAlignMixed         =-2        
	msoBaselineAlignBaseline      =1         
	msoBaselineAlignTop           =2         
	msoBaselineAlignCenter        =3         
	msoBaselineAlignFarEast50     =4         
	msoBaselineAlignAuto          =5         

class MsoBevelType(Enum):
	msoBevelTypeMixed             =-2        
	msoBevelNone                  =1         
	msoBevelRelaxedInset          =2         
	msoBevelCircle                =3         
	msoBevelSlope                 =4         
	msoBevelCross                 =5         
	msoBevelAngle                 =6         
	msoBevelSoftRound             =7         
	msoBevelConvex                =8         
	msoBevelCoolSlant             =9         
	msoBevelDivot                 =10        
	msoBevelRiblet                =11        
	msoBevelHardEdge              =12        
	msoBevelArtDeco               =13        

class MsoBlackWhiteMode(Enum):
	msoBlackWhiteMixed            =-2        
	msoBlackWhiteAutomatic        =1         
	msoBlackWhiteGrayScale        =2         
	msoBlackWhiteLightGrayScale   =3         
	msoBlackWhiteInverseGrayScale =4         
	msoBlackWhiteGrayOutline      =5         
	msoBlackWhiteBlackTextAndLine =6         
	msoBlackWhiteHighContrast     =7         
	msoBlackWhiteBlack            =8         
	msoBlackWhiteWhite            =9         
	msoBlackWhiteDontShow         =10        

class MsoBlogCategorySupport(Enum):
	msoBlogNoCategories           =0         
	msoBlogOneCategory            =1         
	msoBlogMultipleCategories     =2         

class MsoBlogImageType(Enum):
	msoblogImageTypeJPEG          =1         
	msoblogImageTypeGIF           =2         
	msoblogImageTypePNG           =3         

class MsoBroadcastCapabilities(Enum):
	BroadcastCapFileSizeLimited   =1         
	BroadcastCapSupportsMeetingNotes=2         
	BroadcastCapSupportsUpdateDoc =4         

class MsoBroadcastState(Enum):
	NoBroadcast                   =0         
	BroadcastStarted              =1         
	BroadcastPaused               =2         

class MsoBulletType(Enum):
	msoBulletMixed                =-2        
	msoBulletNone                 =0         
	msoBulletUnnumbered           =1         
	msoBulletNumbered             =2         
	msoBulletPicture              =3         

class MsoButtonSetType(Enum):
	msoButtonSetNone              =0         
	msoButtonSetOK                =1         
	msoButtonSetCancel            =2         
	msoButtonSetOkCancel          =3         
	msoButtonSetYesNo             =4         
	msoButtonSetYesNoCancel       =5         
	msoButtonSetBackClose         =6         
	msoButtonSetNextClose         =7         
	msoButtonSetBackNextClose     =8         
	msoButtonSetRetryCancel       =9         
	msoButtonSetAbortRetryIgnore  =10        
	msoButtonSetSearchClose       =11        
	msoButtonSetBackNextSnooze    =12        
	msoButtonSetTipsOptionsClose  =13        
	msoButtonSetYesAllNoCancel    =14        

class MsoButtonState(Enum):
	msoButtonUp                   =0         
	msoButtonDown                 =-1        
	msoButtonMixed                =2         

class MsoButtonStyle(Enum):
	msoButtonAutomatic            =0         
	msoButtonIcon                 =1         
	msoButtonCaption              =2         
	msoButtonIconAndCaption       =3         
	msoButtonIconAndWrapCaption   =7         
	msoButtonIconAndCaptionBelow  =11        
	msoButtonWrapCaption          =14        
	msoButtonIconAndWrapCaptionBelow=15        

class MsoButtonStyleHidden(Enum):
	msoButtonWrapText             =4         
	msoButtonTextBelow            =8         

class MsoCTPDockPosition(Enum):
	msoCTPDockPositionLeft        =0         
	msoCTPDockPositionTop         =1         
	msoCTPDockPositionRight       =2         
	msoCTPDockPositionBottom      =3         
	msoCTPDockPositionFloating    =4         

class MsoCTPDockPositionRestrict(Enum):
	msoCTPDockPositionRestrictNone=0         
	msoCTPDockPositionRestrictNoChange=1         
	msoCTPDockPositionRestrictNoHorizontal=2         
	msoCTPDockPositionRestrictNoVertical=3         

class MsoCalloutAngleType(Enum):
	msoCalloutAngleMixed          =-2        
	msoCalloutAngleAutomatic      =1         
	msoCalloutAngle30             =2         
	msoCalloutAngle45             =3         
	msoCalloutAngle60             =4         
	msoCalloutAngle90             =5         

class MsoCalloutDropType(Enum):
	msoCalloutDropMixed           =-2        
	msoCalloutDropCustom          =1         
	msoCalloutDropTop             =2         
	msoCalloutDropCenter          =3         
	msoCalloutDropBottom          =4         

class MsoCalloutType(Enum):
	msoCalloutMixed               =-2        
	msoCalloutOne                 =1         
	msoCalloutTwo                 =2         
	msoCalloutThree               =3         
	msoCalloutFour                =4         

class MsoCharacterSet(Enum):
	msoCharacterSetArabic         =1         
	msoCharacterSetCyrillic       =2         
	msoCharacterSetEnglishWesternEuropeanOtherLatinScript=3         
	msoCharacterSetGreek          =4         
	msoCharacterSetHebrew         =5         
	msoCharacterSetJapanese       =6         
	msoCharacterSetKorean         =7         
	msoCharacterSetMultilingualUnicode=8         
	msoCharacterSetSimplifiedChinese=9         
	msoCharacterSetThai           =10        
	msoCharacterSetTraditionalChinese=11        
	msoCharacterSetVietnamese     =12        

class MsoChartElementType(Enum):
	msoElementChartTitleNone      =0         
	msoElementChartTitleCenteredOverlay=1         
	msoElementChartTitleAboveChart=2         
	msoElementLegendNone          =100       
	msoElementLegendRight         =101       
	msoElementLegendTop           =102       
	msoElementLegendLeft          =103       
	msoElementLegendBottom        =104       
	msoElementLegendRightOverlay  =105       
	msoElementLegendLeftOverlay   =106       
	msoElementDataLabelNone       =200       
	msoElementDataLabelShow       =201       
	msoElementDataLabelCenter     =202       
	msoElementDataLabelInsideEnd  =203       
	msoElementDataLabelInsideBase =204       
	msoElementDataLabelOutSideEnd =205       
	msoElementDataLabelLeft       =206       
	msoElementDataLabelRight      =207       
	msoElementDataLabelTop        =208       
	msoElementDataLabelBottom     =209       
	msoElementDataLabelBestFit    =210       
	msoElementDataLabelCallout    =211       
	msoElementPrimaryCategoryAxisTitleNone=300       
	msoElementPrimaryCategoryAxisTitleAdjacentToAxis=301       
	msoElementPrimaryCategoryAxisTitleBelowAxis=302       
	msoElementPrimaryCategoryAxisTitleRotated=303       
	msoElementPrimaryCategoryAxisTitleVertical=304       
	msoElementPrimaryCategoryAxisTitleHorizontal=305       
	msoElementPrimaryValueAxisTitleNone=306       
	msoElementPrimaryValueAxisTitleAdjacentToAxis=306       
	msoElementPrimaryValueAxisTitleBelowAxis=308       
	msoElementPrimaryValueAxisTitleRotated=309       
	msoElementPrimaryValueAxisTitleVertical=310       
	msoElementPrimaryValueAxisTitleHorizontal=311       
	msoElementSecondaryCategoryAxisTitleNone=312       
	msoElementSecondaryCategoryAxisTitleAdjacentToAxis=313       
	msoElementSecondaryCategoryAxisTitleBelowAxis=314       
	msoElementSecondaryCategoryAxisTitleRotated=315       
	msoElementSecondaryCategoryAxisTitleVertical=316       
	msoElementSecondaryCategoryAxisTitleHorizontal=317       
	msoElementSecondaryValueAxisTitleNone=318       
	msoElementSecondaryValueAxisTitleAdjacentToAxis=319       
	msoElementSecondaryValueAxisTitleBelowAxis=320       
	msoElementSecondaryValueAxisTitleRotated=321       
	msoElementSecondaryValueAxisTitleVertical=322       
	msoElementSecondaryValueAxisTitleHorizontal=323       
	msoElementSeriesAxisTitleNone =324       
	msoElementSeriesAxisTitleRotated=325       
	msoElementSeriesAxisTitleVertical=326       
	msoElementSeriesAxisTitleHorizontal=327       
	msoElementPrimaryValueGridLinesNone=328       
	msoElementPrimaryValueGridLinesMinor=329       
	msoElementPrimaryValueGridLinesMajor=330       
	msoElementPrimaryValueGridLinesMinorMajor=331       
	msoElementPrimaryCategoryGridLinesNone=332       
	msoElementPrimaryCategoryGridLinesMinor=333       
	msoElementPrimaryCategoryGridLinesMajor=334       
	msoElementPrimaryCategoryGridLinesMinorMajor=335       
	msoElementSecondaryValueGridLinesNone=336       
	msoElementSecondaryValueGridLinesMinor=337       
	msoElementSecondaryValueGridLinesMajor=338       
	msoElementSecondaryValueGridLinesMinorMajor=339       
	msoElementSecondaryCategoryGridLinesNone=340       
	msoElementSecondaryCategoryGridLinesMinor=341       
	msoElementSecondaryCategoryGridLinesMajor=342       
	msoElementSecondaryCategoryGridLinesMinorMajor=343       
	msoElementSeriesAxisGridLinesNone=344       
	msoElementSeriesAxisGridLinesMinor=345       
	msoElementSeriesAxisGridLinesMajor=346       
	msoElementSeriesAxisGridLinesMinorMajor=347       
	msoElementPrimaryCategoryAxisNone=348       
	msoElementPrimaryCategoryAxisShow=349       
	msoElementPrimaryCategoryAxisWithoutLabels=350       
	msoElementPrimaryCategoryAxisReverse=351       
	msoElementPrimaryValueAxisNone=352       
	msoElementPrimaryValueAxisShow=353       
	msoElementPrimaryValueAxisThousands=354       
	msoElementPrimaryValueAxisMillions=355       
	msoElementPrimaryValueAxisBillions=356       
	msoElementPrimaryValueAxisLogScale=357       
	msoElementSecondaryCategoryAxisNone=358       
	msoElementSecondaryCategoryAxisShow=359       
	msoElementSecondaryCategoryAxisWithoutLabels=360       
	msoElementSecondaryCategoryAxisReverse=361       
	msoElementSecondaryValueAxisNone=362       
	msoElementSecondaryValueAxisShow=363       
	msoElementSecondaryValueAxisThousands=364       
	msoElementSecondaryValueAxisMillions=365       
	msoElementSecondaryValueAxisBillions=366       
	msoElementSecondaryValueAxisLogScale=367       
	msoElementSeriesAxisNone      =368       
	msoElementSeriesAxisShow      =369       
	msoElementSeriesAxisWithoutLabeling=370       
	msoElementSeriesAxisReverse   =371       
	msoElementPrimaryCategoryAxisThousands=372       
	msoElementPrimaryCategoryAxisMillions=373       
	msoElementPrimaryCategoryAxisBillions=374       
	msoElementPrimaryCategoryAxisLogScale=375       
	msoElementSecondaryCategoryAxisThousands=376       
	msoElementSecondaryCategoryAxisMillions=377       
	msoElementSecondaryCategoryAxisBillions=378       
	msoElementSecondaryCategoryAxisLogScale=379       
	msoElementDataTableNone       =500       
	msoElementDataTableShow       =501       
	msoElementDataTableWithLegendKeys=502       
	msoElementTrendlineNone       =600       
	msoElementTrendlineAddLinear  =601       
	msoElementTrendlineAddExponential=602       
	msoElementTrendlineAddLinearForecast=603       
	msoElementTrendlineAddTwoPeriodMovingAverage=604       
	msoElementErrorBarNone        =700       
	msoElementErrorBarStandardError=701       
	msoElementErrorBarPercentage  =702       
	msoElementErrorBarStandardDeviation=703       
	msoElementLineNone            =800       
	msoElementLineDropLine        =801       
	msoElementLineHiLoLine        =802       
	msoElementLineSeriesLine      =803       
	msoElementLineDropHiLoLine    =804       
	msoElementUpDownBarsNone      =900       
	msoElementUpDownBarsShow      =901       
	msoElementPlotAreaNone        =1000      
	msoElementPlotAreaShow        =1001      
	msoElementChartWallNone       =1100      
	msoElementChartWallShow       =1101      
	msoElementChartFloorNone      =1200      
	msoElementChartFloorShow      =1201      

class MsoChartFieldType(Enum):
	msoChartFieldBubbleSize       =1         
	msoChartFieldCategoryName     =2         
	msoChartFieldPercentage       =3         
	msoChartFieldSeriesName       =4         
	msoChartFieldValue            =5         
	msoChartFieldFormula          =6         
	msoChartFieldRange            =7         

class MsoClipboardFormat(Enum):
	msoClipboardFormatMixed       =-2        
	msoClipboardFormatNative      =1         
	msoClipboardFormatHTML        =2         
	msoClipboardFormatRTF         =3         
	msoClipboardFormatPlainText   =4         

class MsoColorType(Enum):
	msoColorTypeMixed             =-2        
	msoColorTypeRGB               =1         
	msoColorTypeScheme            =2         
	msoColorTypeCMYK              =3         
	msoColorTypeCMS               =4         
	msoColorTypeInk               =5         

class MsoComboStyle(Enum):
	msoComboNormal                =0         
	msoComboLabel                 =1         

class MsoCommandBarButtonHyperlinkType(Enum):
	msoCommandBarButtonHyperlinkNone=0         
	msoCommandBarButtonHyperlinkOpen=1         
	msoCommandBarButtonHyperlinkInsertPicture=2         

class MsoCondition(Enum):
	msoConditionFileTypeAllFiles  =1         
	msoConditionFileTypeOfficeFiles=2         
	msoConditionFileTypeWordDocuments=3         
	msoConditionFileTypeExcelWorkbooks=4         
	msoConditionFileTypePowerPointPresentations=5         
	msoConditionFileTypeBinders   =6         
	msoConditionFileTypeDatabases =7         
	msoConditionFileTypeTemplates =8         
	msoConditionIncludes          =9         
	msoConditionIncludesPhrase    =10        
	msoConditionBeginsWith        =11        
	msoConditionEndsWith          =12        
	msoConditionIncludesNearEachOther=13        
	msoConditionIsExactly         =14        
	msoConditionIsNot             =15        
	msoConditionYesterday         =16        
	msoConditionToday             =17        
	msoConditionTomorrow          =18        
	msoConditionLastWeek          =19        
	msoConditionThisWeek          =20        
	msoConditionNextWeek          =21        
	msoConditionLastMonth         =22        
	msoConditionThisMonth         =23        
	msoConditionNextMonth         =24        
	msoConditionAnytime           =25        
	msoConditionAnytimeBetween    =26        
	msoConditionOn                =27        
	msoConditionOnOrAfter         =28        
	msoConditionOnOrBefore        =29        
	msoConditionInTheNext         =30        
	msoConditionInTheLast         =31        
	msoConditionEquals            =32        
	msoConditionDoesNotEqual      =33        
	msoConditionAnyNumberBetween  =34        
	msoConditionAtMost            =35        
	msoConditionAtLeast           =36        
	msoConditionMoreThan          =37        
	msoConditionLessThan          =38        
	msoConditionIsYes             =39        
	msoConditionIsNo              =40        
	msoConditionIncludesFormsOf   =41        
	msoConditionFreeText          =42        
	msoConditionFileTypeOutlookItems=43        
	msoConditionFileTypeMailItem  =44        
	msoConditionFileTypeCalendarItem=45        
	msoConditionFileTypeContactItem=46        
	msoConditionFileTypeNoteItem  =47        
	msoConditionFileTypeJournalItem=48        
	msoConditionFileTypeTaskItem  =49        
	msoConditionFileTypePhotoDrawFiles=50        
	msoConditionFileTypeDataConnectionFiles=51        
	msoConditionFileTypePublisherFiles=52        
	msoConditionFileTypeProjectFiles=53        
	msoConditionFileTypeDocumentImagingFiles=54        
	msoConditionFileTypeVisioFiles=55        
	msoConditionFileTypeDesignerFiles=56        
	msoConditionFileTypeWebPages  =57        
	msoConditionEqualsLow         =58        
	msoConditionEqualsNormal      =59        
	msoConditionEqualsHigh        =60        
	msoConditionNotEqualToLow     =61        
	msoConditionNotEqualToNormal  =62        
	msoConditionNotEqualToHigh    =63        
	msoConditionEqualsNotStarted  =64        
	msoConditionEqualsInProgress  =65        
	msoConditionEqualsCompleted   =66        
	msoConditionEqualsWaitingForSomeoneElse=67        
	msoConditionEqualsDeferred    =68        
	msoConditionNotEqualToNotStarted=69        
	msoConditionNotEqualToInProgress=70        
	msoConditionNotEqualToCompleted=71        
	msoConditionNotEqualToWaitingForSomeoneElse=72        
	msoConditionNotEqualToDeferred=73        

class MsoConnector(Enum):
	msoConnectorAnd               =1         
	msoConnectorOr                =2         

class MsoConnectorType(Enum):
	msoConnectorTypeMixed         =-2        
	msoConnectorStraight          =1         
	msoConnectorElbow             =2         
	msoConnectorCurve             =3         

class MsoContactCardAddressType(Enum):
	msoContactCardAddressTypeUnknown=0         
	msoContactCardAddressTypeOutlook=1         
	msoContactCardAddressTypeSMTP =2         
	msoContactCardAddressTypeIM   =3         

class MsoContactCardStyle(Enum):
	msoContactCardHover           =0         
	msoContactCardFull            =1         

class MsoContactCardType(Enum):
	msoContactCardTypeEnterpriseContact=0         
	msoContactCardTypePersonalContact=1         
	msoContactCardTypeUnknownContact=2         
	msoContactCardTypeEnterpriseGroup=3         
	msoContactCardTypePersonalDistributionList=4         

class MsoControlOLEUsage(Enum):
	msoControlOLEUsageNeither     =0         
	msoControlOLEUsageServer      =1         
	msoControlOLEUsageClient      =2         
	msoControlOLEUsageBoth        =3         

class MsoControlType(Enum):
	msoControlCustom              =0         
	msoControlButton              =1         
	msoControlEdit                =2         
	msoControlDropdown            =3         
	msoControlComboBox            =4         
	msoControlButtonDropdown      =5         
	msoControlSplitDropdown       =6         
	msoControlOCXDropdown         =7         
	msoControlGenericDropdown     =8         
	msoControlGraphicDropdown     =9         
	msoControlPopup               =10        
	msoControlGraphicPopup        =11        
	msoControlButtonPopup         =12        
	msoControlSplitButtonPopup    =13        
	msoControlSplitButtonMRUPopup =14        
	msoControlLabel               =15        
	msoControlExpandingGrid       =16        
	msoControlSplitExpandingGrid  =17        
	msoControlGrid                =18        
	msoControlGauge               =19        
	msoControlGraphicCombo        =20        
	msoControlPane                =21        
	msoControlActiveX             =22        
	msoControlSpinner             =23        
	msoControlLabelEx             =24        
	msoControlWorkPane            =25        
	msoControlAutoCompleteCombo   =26        

class MsoCustomXMLNodeType(Enum):
	msoCustomXMLNodeElement       =1         
	msoCustomXMLNodeAttribute     =2         
	msoCustomXMLNodeText          =3         
	msoCustomXMLNodeCData         =4         
	msoCustomXMLNodeProcessingInstruction=7         
	msoCustomXMLNodeComment       =8         
	msoCustomXMLNodeDocument      =9         

class MsoCustomXMLValidationErrorType(Enum):
	msoCustomXMLValidationErrorSchemaGenerated=0         
	msoCustomXMLValidationErrorAutomaticallyCleared=1         
	msoCustomXMLValidationErrorManual=2         

class MsoDateTimeFormat(Enum):
	msoDateTimeFormatMixed        =-2        
	msoDateTimeMdyy               =1         
	msoDateTimeddddMMMMddyyyy     =2         
	msoDateTimedMMMMyyyy          =3         
	msoDateTimeMMMMdyyyy          =4         
	msoDateTimedMMMyy             =5         
	msoDateTimeMMMMyy             =6         
	msoDateTimeMMyy               =7         
	msoDateTimeMMddyyHmm          =8         
	msoDateTimeMMddyyhmmAMPM      =9         
	msoDateTimeHmm                =10        
	msoDateTimeHmmss              =11        
	msoDateTimehmmAMPM            =12        
	msoDateTimehmmssAMPM          =13        
	msoDateTimeFigureOut          =14        

class MsoDiagramNodeType(Enum):
	msoDiagramNode                =1         
	msoDiagramAssistant           =2         

class MsoDiagramType(Enum):
	msoDiagramMixed               =-2        
	msoDiagramOrgChart            =1         
	msoDiagramCycle               =2         
	msoDiagramRadial              =3         
	msoDiagramPyramid             =4         
	msoDiagramVenn                =5         
	msoDiagramTarget              =6         

class MsoDistributeCmd(Enum):
	msoDistributeHorizontally     =0         
	msoDistributeVertically       =1         

class MsoDocInspectorStatus(Enum):
	msoDocInspectorStatusDocOk    =0         
	msoDocInspectorStatusIssueFound=1         
	msoDocInspectorStatusError    =2         

class MsoDocProperties(Enum):
	msoPropertyTypeNumber         =1         
	msoPropertyTypeBoolean        =2         
	msoPropertyTypeDate           =3         
	msoPropertyTypeString         =4         
	msoPropertyTypeFloat          =5         

class MsoEditingType(Enum):
	msoEditingAuto                =0         
	msoEditingCorner              =1         
	msoEditingSmooth              =2         
	msoEditingSymmetric           =3         

class MsoEncoding(Enum):
	msoEncodingThai               =874       
	msoEncodingJapaneseShiftJIS   =932       
	msoEncodingSimplifiedChineseGBK=936       
	msoEncodingKorean             =949       
	msoEncodingTraditionalChineseBig5=950       
	msoEncodingUnicodeLittleEndian=1200      
	msoEncodingUnicodeBigEndian   =1201      
	msoEncodingCentralEuropean    =1250      
	msoEncodingCyrillic           =1251      
	msoEncodingWestern            =1252      
	msoEncodingGreek              =1253      
	msoEncodingTurkish            =1254      
	msoEncodingHebrew             =1255      
	msoEncodingArabic             =1256      
	msoEncodingBaltic             =1257      
	msoEncodingVietnamese         =1258      
	msoEncodingAutoDetect         =50001     
	msoEncodingJapaneseAutoDetect =50932     
	msoEncodingSimplifiedChineseAutoDetect=50936     
	msoEncodingKoreanAutoDetect   =50949     
	msoEncodingTraditionalChineseAutoDetect=50950     
	msoEncodingCyrillicAutoDetect =51251     
	msoEncodingGreekAutoDetect    =51253     
	msoEncodingArabicAutoDetect   =51256     
	msoEncodingISO88591Latin1     =28591     
	msoEncodingISO88592CentralEurope=28592     
	msoEncodingISO88593Latin3     =28593     
	msoEncodingISO88594Baltic     =28594     
	msoEncodingISO88595Cyrillic   =28595     
	msoEncodingISO88596Arabic     =28596     
	msoEncodingISO88597Greek      =28597     
	msoEncodingISO88598Hebrew     =28598     
	msoEncodingISO88599Turkish    =28599     
	msoEncodingISO885915Latin9    =28605     
	msoEncodingISO88598HebrewLogical=38598     
	msoEncodingISO2022JPNoHalfwidthKatakana=50220     
	msoEncodingISO2022JPJISX02021984=50221     
	msoEncodingISO2022JPJISX02011989=50222     
	msoEncodingISO2022KR          =50225     
	msoEncodingISO2022CNTraditionalChinese=50227     
	msoEncodingISO2022CNSimplifiedChinese=50229     
	msoEncodingMacRoman           =10000     
	msoEncodingMacJapanese        =10001     
	msoEncodingMacTraditionalChineseBig5=10002     
	msoEncodingMacKorean          =10003     
	msoEncodingMacArabic          =10004     
	msoEncodingMacHebrew          =10005     
	msoEncodingMacGreek1          =10006     
	msoEncodingMacCyrillic        =10007     
	msoEncodingMacSimplifiedChineseGB2312=10008     
	msoEncodingMacRomania         =10010     
	msoEncodingMacUkraine         =10017     
	msoEncodingMacLatin2          =10029     
	msoEncodingMacIcelandic       =10079     
	msoEncodingMacTurkish         =10081     
	msoEncodingMacCroatia         =10082     
	msoEncodingEBCDICUSCanada     =37        
	msoEncodingEBCDICInternational=500       
	msoEncodingEBCDICMultilingualROECELatin2=870       
	msoEncodingEBCDICGreekModern  =875       
	msoEncodingEBCDICTurkishLatin5=1026      
	msoEncodingEBCDICGermany      =20273     
	msoEncodingEBCDICDenmarkNorway=20277     
	msoEncodingEBCDICFinlandSweden=20278     
	msoEncodingEBCDICItaly        =20280     
	msoEncodingEBCDICLatinAmericaSpain=20284     
	msoEncodingEBCDICUnitedKingdom=20285     
	msoEncodingEBCDICJapaneseKatakanaExtended=20290     
	msoEncodingEBCDICFrance       =20297     
	msoEncodingEBCDICArabic       =20420     
	msoEncodingEBCDICGreek        =20423     
	msoEncodingEBCDICHebrew       =20424     
	msoEncodingEBCDICKoreanExtended=20833     
	msoEncodingEBCDICThai         =20838     
	msoEncodingEBCDICIcelandic    =20871     
	msoEncodingEBCDICTurkish      =20905     
	msoEncodingEBCDICRussian      =20880     
	msoEncodingEBCDICSerbianBulgarian=21025     
	msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese=50930     
	msoEncodingEBCDICUSCanadaAndJapanese=50931     
	msoEncodingEBCDICKoreanExtendedAndKorean=50933     
	msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese=50935     
	msoEncodingEBCDICUSCanadaAndTraditionalChinese=50937     
	msoEncodingEBCDICJapaneseLatinExtendedAndJapanese=50939     
	msoEncodingOEMUnitedStates    =437       
	msoEncodingOEMGreek437G       =737       
	msoEncodingOEMBaltic          =775       
	msoEncodingOEMMultilingualLatinI=850       
	msoEncodingOEMMultilingualLatinII=852       
	msoEncodingOEMCyrillic        =855       
	msoEncodingOEMTurkish         =857       
	msoEncodingOEMPortuguese      =860       
	msoEncodingOEMIcelandic       =861       
	msoEncodingOEMHebrew          =862       
	msoEncodingOEMCanadianFrench  =863       
	msoEncodingOEMArabic          =864       
	msoEncodingOEMNordic          =865       
	msoEncodingOEMCyrillicII      =866       
	msoEncodingOEMModernGreek     =869       
	msoEncodingEUCJapanese        =51932     
	msoEncodingEUCChineseSimplifiedChinese=51936     
	msoEncodingEUCKorean          =51949     
	msoEncodingEUCTaiwaneseTraditionalChinese=51950     
	msoEncodingISCIIDevanagari    =57002     
	msoEncodingISCIIBengali       =57003     
	msoEncodingISCIITamil         =57004     
	msoEncodingISCIITelugu        =57005     
	msoEncodingISCIIAssamese      =57006     
	msoEncodingISCIIOriya         =57007     
	msoEncodingISCIIKannada       =57008     
	msoEncodingISCIIMalayalam     =57009     
	msoEncodingISCIIGujarati      =57010     
	msoEncodingISCIIPunjabi       =57011     
	msoEncodingArabicASMO         =708       
	msoEncodingArabicTransparentASMO=720       
	msoEncodingKoreanJohab        =1361      
	msoEncodingTaiwanCNS          =20000     
	msoEncodingTaiwanTCA          =20001     
	msoEncodingTaiwanEten         =20002     
	msoEncodingTaiwanIBM5550      =20003     
	msoEncodingTaiwanTeleText     =20004     
	msoEncodingTaiwanWang         =20005     
	msoEncodingIA5IRV             =20105     
	msoEncodingIA5German          =20106     
	msoEncodingIA5Swedish         =20107     
	msoEncodingIA5Norwegian       =20108     
	msoEncodingUSASCII            =20127     
	msoEncodingT61                =20261     
	msoEncodingISO6937NonSpacingAccent=20269     
	msoEncodingKOI8R              =20866     
	msoEncodingExtAlphaLowercase  =21027     
	msoEncodingKOI8U              =21866     
	msoEncodingEuropa3            =29001     
	msoEncodingHZGBSimplifiedChinese=52936     
	msoEncodingSimplifiedChineseGB18030=54936     
	msoEncodingUTF7               =65000     
	msoEncodingUTF8               =65001     

class MsoExtraInfoMethod(Enum):
	msoMethodGet                  =0         
	msoMethodPost                 =1         

class MsoExtrusionColorType(Enum):
	msoExtrusionColorTypeMixed    =-2        
	msoExtrusionColorAutomatic    =1         
	msoExtrusionColorCustom       =2         

class MsoFarEastLineBreakLanguageID(Enum):
	MsoFarEastLineBreakLanguageJapanese=1041      
	MsoFarEastLineBreakLanguageKorean=1042      
	MsoFarEastLineBreakLanguageSimplifiedChinese=2052      
	MsoFarEastLineBreakLanguageTraditionalChinese=1028      

class MsoFeatureInstall(Enum):
	msoFeatureInstallNone         =0         
	msoFeatureInstallOnDemand     =1         
	msoFeatureInstallOnDemandWithUI=2         

class MsoFileDialogType(Enum):
	msoFileDialogOpen             =1         
	msoFileDialogSaveAs           =2         
	msoFileDialogFilePicker       =3         
	msoFileDialogFolderPicker     =4         

class MsoFileDialogView(Enum):
	msoFileDialogViewList         =1         
	msoFileDialogViewDetails      =2         
	msoFileDialogViewProperties   =3         
	msoFileDialogViewPreview      =4         
	msoFileDialogViewThumbnail    =5         
	msoFileDialogViewLargeIcons   =6         
	msoFileDialogViewSmallIcons   =7         
	msoFileDialogViewWebView      =8         
	msoFileDialogViewTiles        =9         

class MsoFileFindListBy(Enum):
	msoListbyName                 =1         
	msoListbyTitle                =2         

class MsoFileFindOptions(Enum):
	msoOptionsNew                 =1         
	msoOptionsAdd                 =2         
	msoOptionsWithin              =3         

class MsoFileFindSortBy(Enum):
	msoFileFindSortbyAuthor       =1         
	msoFileFindSortbyDateCreated  =2         
	msoFileFindSortbyLastSavedBy  =3         
	msoFileFindSortbyDateSaved    =4         
	msoFileFindSortbyFileName     =5         
	msoFileFindSortbySize         =6         
	msoFileFindSortbyTitle        =7         

class MsoFileFindView(Enum):
	msoViewFileInfo               =1         
	msoViewPreview                =2         
	msoViewSummaryInfo            =3         

class MsoFileNewAction(Enum):
	msoEditFile                   =0         
	msoCreateNewFile              =1         
	msoOpenFile                   =2         

class MsoFileNewSection(Enum):
	msoOpenDocument               =0         
	msoNew                        =1         
	msoNewfromExistingFile        =2         
	msoNewfromTemplate            =3         
	msoBottomSection              =4         

class MsoFileType(Enum):
	msoFileTypeAllFiles           =1         
	msoFileTypeOfficeFiles        =2         
	msoFileTypeWordDocuments      =3         
	msoFileTypeExcelWorkbooks     =4         
	msoFileTypePowerPointPresentations=5         
	msoFileTypeBinders            =6         
	msoFileTypeDatabases          =7         
	msoFileTypeTemplates          =8         
	msoFileTypeOutlookItems       =9         
	msoFileTypeMailItem           =10        
	msoFileTypeCalendarItem       =11        
	msoFileTypeContactItem        =12        
	msoFileTypeNoteItem           =13        
	msoFileTypeJournalItem        =14        
	msoFileTypeTaskItem           =15        
	msoFileTypePhotoDrawFiles     =16        
	msoFileTypeDataConnectionFiles=17        
	msoFileTypePublisherFiles     =18        
	msoFileTypeProjectFiles       =19        
	msoFileTypeDocumentImagingFiles=20        
	msoFileTypeVisioFiles         =21        
	msoFileTypeDesignerFiles      =22        
	msoFileTypeWebPages           =23        

class MsoFileValidationMode(Enum):
	msoFileValidationDefault      =0         
	msoFileValidationSkip         =1         

class MsoFillType(Enum):
	msoFillMixed                  =-2        
	msoFillSolid                  =1         
	msoFillPatterned              =2         
	msoFillGradient               =3         
	msoFillTextured               =4         
	msoFillBackground             =5         
	msoFillPicture                =6         

class MsoFilterComparison(Enum):
	msoFilterComparisonEqual      =0         
	msoFilterComparisonNotEqual   =1         
	msoFilterComparisonLessThan   =2         
	msoFilterComparisonGreaterThan=3         
	msoFilterComparisonLessThanEqual=4         
	msoFilterComparisonGreaterThanEqual=5         
	msoFilterComparisonIsBlank    =6         
	msoFilterComparisonIsNotBlank =7         
	msoFilterComparisonContains   =8         
	msoFilterComparisonNotContains=9         

class MsoFilterConjunction(Enum):
	msoFilterConjunctionAnd       =0         
	msoFilterConjunctionOr        =1         

class MsoFlipCmd(Enum):
	msoFlipHorizontal             =0         
	msoFlipVertical               =1         

class MsoFontLanguageIndex(Enum):
	msoThemeLatin                 =1         
	msoThemeComplexScript         =2         
	msoThemeEastAsian             =3         

class MsoGradientColorType(Enum):
	msoGradientColorMixed         =-2        
	msoGradientOneColor           =1         
	msoGradientTwoColors          =2         
	msoGradientPresetColors       =3         
	msoGradientMultiColor         =4         

class MsoGradientStyle(Enum):
	msoGradientMixed              =-2        
	msoGradientHorizontal         =1         
	msoGradientVertical           =2         
	msoGradientDiagonalUp         =3         
	msoGradientDiagonalDown       =4         
	msoGradientFromCorner         =5         
	msoGradientFromTitle          =6         
	msoGradientFromCenter         =7         

class MsoGraphicStyleIndex(Enum):
	msoGraphicStyleMixed          =-2        
	msoGraphicStyleNotAPreset     =0         
	msoGraphicStylePreset1        =1         
	msoGraphicStylePreset2        =2         
	msoGraphicStylePreset3        =3         
	msoGraphicStylePreset4        =4         
	msoGraphicStylePreset5        =5         
	msoGraphicStylePreset6        =6         
	msoGraphicStylePreset7        =7         
	msoGraphicStylePreset8        =8         
	msoGraphicStylePreset9        =9         
	msoGraphicStylePreset10       =10        
	msoGraphicStylePreset11       =11        
	msoGraphicStylePreset12       =12        
	msoGraphicStylePreset13       =13        
	msoGraphicStylePreset14       =14        
	msoGraphicStylePreset15       =15        
	msoGraphicStylePreset16       =16        
	msoGraphicStylePreset17       =17        
	msoGraphicStylePreset18       =18        
	msoGraphicStylePreset19       =19        
	msoGraphicStylePreset20       =20        
	msoGraphicStylePreset21       =21        
	msoGraphicStylePreset22       =22        
	msoGraphicStylePreset23       =23        
	msoGraphicStylePreset24       =24        
	msoGraphicStylePreset25       =25        
	msoGraphicStylePreset26       =26        
	msoGraphicStylePreset27       =27        
	msoGraphicStylePreset28       =28        

class MsoHTMLProjectOpen(Enum):
	msoHTMLProjectOpenSourceView  =1         
	msoHTMLProjectOpenTextView    =2         

class MsoHTMLProjectState(Enum):
	msoHTMLProjectStateDocumentLocked=1         
	msoHTMLProjectStateProjectLocked=2         
	msoHTMLProjectStateDocumentProjectUnlocked=3         

class MsoHorizontalAnchor(Enum):
	msoHorizontalAnchorMixed      =-2        
	msoAnchorNone                 =1         
	msoAnchorCenter               =2         

class MsoHyperlinkType(Enum):
	msoHyperlinkRange             =0         
	msoHyperlinkShape             =1         
	msoHyperlinkInlineShape       =2         

class MsoIconType(Enum):
	msoIconNone                   =0         
	msoIconAlert                  =2         
	msoIconTip                    =3         
	msoIconAlertInfo              =4         
	msoIconAlertWarning           =5         
	msoIconAlertQuery             =6         
	msoIconAlertCritical          =7         

class MsoIodGroup(Enum):
	msoIodGroupPIAs               =0         
	msoIodGroupVSTOR35Mgd         =1         
	msoIodGroupVSTOR40Mgd         =2         

class MsoLanguageID(Enum):
	msoLanguageIDMixed            =-2        
	msoLanguageIDNone             =0         
	msoLanguageIDNoProofing       =1024      
	msoLanguageIDAfrikaans        =1078      
	msoLanguageIDAlbanian         =1052      
	msoLanguageIDAmharic          =1118      
	msoLanguageIDArabicAlgeria    =5121      
	msoLanguageIDArabicBahrain    =15361     
	msoLanguageIDArabicEgypt      =3073      
	msoLanguageIDArabicIraq       =2049      
	msoLanguageIDArabicJordan     =11265     
	msoLanguageIDArabicKuwait     =13313     
	msoLanguageIDArabicLebanon    =12289     
	msoLanguageIDArabicLibya      =4097      
	msoLanguageIDArabicMorocco    =6145      
	msoLanguageIDArabicOman       =8193      
	msoLanguageIDArabicQatar      =16385     
	msoLanguageIDArabic           =1025      
	msoLanguageIDArabicSyria      =10241     
	msoLanguageIDArabicTunisia    =7169      
	msoLanguageIDArabicUAE        =14337     
	msoLanguageIDArabicYemen      =9217      
	msoLanguageIDArmenian         =1067      
	msoLanguageIDAssamese         =1101      
	msoLanguageIDAzeriCyrillic    =2092      
	msoLanguageIDAzeriLatin       =1068      
	msoLanguageIDBasque           =1069      
	msoLanguageIDByelorussian     =1059      
	msoLanguageIDBengali          =1093      
	msoLanguageIDBosnian          =4122      
	msoLanguageIDBosnianBosniaHerzegovinaCyrillic=8218      
	msoLanguageIDBosnianBosniaHerzegovinaLatin=5146      
	msoLanguageIDBulgarian        =1026      
	msoLanguageIDBurmese          =1109      
	msoLanguageIDCatalan          =1027      
	msoLanguageIDChineseHongKongSAR=3076      
	msoLanguageIDChineseMacaoSAR  =5124      
	msoLanguageIDSimplifiedChinese=2052      
	msoLanguageIDChineseSingapore =4100      
	msoLanguageIDTraditionalChinese=1028      
	msoLanguageIDCherokee         =1116      
	msoLanguageIDCroatian         =1050      
	msoLanguageIDCzech            =1029      
	msoLanguageIDDanish           =1030      
	msoLanguageIDDivehi           =1125      
	msoLanguageIDBelgianDutch     =2067      
	msoLanguageIDDutch            =1043      
	msoLanguageIDDzongkhaBhutan   =2129      
	msoLanguageIDEdo              =1126      
	msoLanguageIDEnglishAUS       =3081      
	msoLanguageIDEnglishBelize    =10249     
	msoLanguageIDEnglishCanadian  =4105      
	msoLanguageIDEnglishCaribbean =9225      
	msoLanguageIDEnglishIndonesia =14345     
	msoLanguageIDEnglishIreland   =6153      
	msoLanguageIDEnglishJamaica   =8201      
	msoLanguageIDEnglishNewZealand=5129      
	msoLanguageIDEnglishPhilippines=13321     
	msoLanguageIDEnglishSouthAfrica=7177      
	msoLanguageIDEnglishTrinidadTobago=11273     
	msoLanguageIDEnglishUK        =2057      
	msoLanguageIDEnglishUS        =1033      
	msoLanguageIDEnglishZimbabwe  =12297     
	msoLanguageIDEstonian         =1061      
	msoLanguageIDFaeroese         =1080      
	msoLanguageIDFarsi            =1065      
	msoLanguageIDFilipino         =1124      
	msoLanguageIDFinnish          =1035      
	msoLanguageIDBelgianFrench    =2060      
	msoLanguageIDFrenchCameroon   =11276     
	msoLanguageIDFrenchCanadian   =3084      
	msoLanguageIDFrenchCotedIvoire=12300     
	msoLanguageIDFrench           =1036      
	msoLanguageIDFrenchHaiti      =15372     
	msoLanguageIDFrenchLuxembourg =5132      
	msoLanguageIDFrenchMali       =13324     
	msoLanguageIDFrenchMonaco     =6156      
	msoLanguageIDFrenchMorocco    =14348     
	msoLanguageIDFrenchReunion    =8204      
	msoLanguageIDFrenchSenegal    =10252     
	msoLanguageIDSwissFrench      =4108      
	msoLanguageIDFrenchWestIndies =7180      
	msoLanguageIDFrenchZaire      =9228      
	msoLanguageIDFrenchCongoDRC   =9228      
	msoLanguageIDFrisianNetherlands=1122      
	msoLanguageIDFulfulde         =1127      
	msoLanguageIDGaelicIreland    =2108      
	msoLanguageIDGaelicScotland   =1084      
	msoLanguageIDGalician         =1110      
	msoLanguageIDGeorgian         =1079      
	msoLanguageIDGermanAustria    =3079      
	msoLanguageIDGerman           =1031      
	msoLanguageIDGermanLiechtenstein=5127      
	msoLanguageIDGermanLuxembourg =4103      
	msoLanguageIDSwissGerman      =2055      
	msoLanguageIDGreek            =1032      
	msoLanguageIDGuarani          =1140      
	msoLanguageIDGujarati         =1095      
	msoLanguageIDHausa            =1128      
	msoLanguageIDHawaiian         =1141      
	msoLanguageIDHebrew           =1037      
	msoLanguageIDHindi            =1081      
	msoLanguageIDHungarian        =1038      
	msoLanguageIDIbibio           =1129      
	msoLanguageIDIcelandic        =1039      
	msoLanguageIDIgbo             =1136      
	msoLanguageIDIndonesian       =1057      
	msoLanguageIDInuktitut        =1117      
	msoLanguageIDItalian          =1040      
	msoLanguageIDSwissItalian     =2064      
	msoLanguageIDJapanese         =1041      
	msoLanguageIDKannada          =1099      
	msoLanguageIDKanuri           =1137      
	msoLanguageIDKashmiri         =1120      
	msoLanguageIDKashmiriDevanagari=2144      
	msoLanguageIDKazakh           =1087      
	msoLanguageIDKhmer            =1107      
	msoLanguageIDKirghiz          =1088      
	msoLanguageIDKonkani          =1111      
	msoLanguageIDKorean           =1042      
	msoLanguageIDKyrgyz           =1088      
	msoLanguageIDLatin            =1142      
	msoLanguageIDLao              =1108      
	msoLanguageIDLatvian          =1062      
	msoLanguageIDLithuanian       =1063      
	msoLanguageIDMacedonian       =1071      
	msoLanguageIDMacedonianFYROM  =1071      
	msoLanguageIDMalaysian        =1086      
	msoLanguageIDMalayBruneiDarussalam=2110      
	msoLanguageIDMalayalam        =1100      
	msoLanguageIDMaltese          =1082      
	msoLanguageIDManipuri         =1112      
	msoLanguageIDMaori            =1153      
	msoLanguageIDMarathi          =1102      
	msoLanguageIDMongolian        =1104      
	msoLanguageIDNepali           =1121      
	msoLanguageIDNorwegianBokmol  =1044      
	msoLanguageIDNorwegianNynorsk =2068      
	msoLanguageIDOriya            =1096      
	msoLanguageIDOromo            =1138      
	msoLanguageIDPashto           =1123      
	msoLanguageIDPolish           =1045      
	msoLanguageIDBrazilianPortuguese=1046      
	msoLanguageIDPortuguese       =2070      
	msoLanguageIDPunjabi          =1094      
	msoLanguageIDQuechuaBolivia   =1131      
	msoLanguageIDQuechuaEcuador   =2155      
	msoLanguageIDQuechuaPeru      =3179      
	msoLanguageIDRhaetoRomanic    =1047      
	msoLanguageIDRomanianMoldova  =2072      
	msoLanguageIDRomanian         =1048      
	msoLanguageIDRussianMoldova   =2073      
	msoLanguageIDRussian          =1049      
	msoLanguageIDSamiLappish      =1083      
	msoLanguageIDSanskrit         =1103      
	msoLanguageIDSepedi           =1132      
	msoLanguageIDSerbianBosniaHerzegovinaCyrillic=7194      
	msoLanguageIDSerbianBosniaHerzegovinaLatin=6170      
	msoLanguageIDSerbianCyrillic  =3098      
	msoLanguageIDSerbianLatin     =2074      
	msoLanguageIDSesotho          =1072      
	msoLanguageIDSindhi           =1113      
	msoLanguageIDSindhiPakistan   =2137      
	msoLanguageIDSinhalese        =1115      
	msoLanguageIDSlovak           =1051      
	msoLanguageIDSlovenian        =1060      
	msoLanguageIDSomali           =1143      
	msoLanguageIDSorbian          =1070      
	msoLanguageIDSpanishArgentina =11274     
	msoLanguageIDSpanishBolivia   =16394     
	msoLanguageIDSpanishChile     =13322     
	msoLanguageIDSpanishColombia  =9226      
	msoLanguageIDSpanishCostaRica =5130      
	msoLanguageIDSpanishDominicanRepublic=7178      
	msoLanguageIDSpanishEcuador   =12298     
	msoLanguageIDSpanishElSalvador=17418     
	msoLanguageIDSpanishGuatemala =4106      
	msoLanguageIDSpanishHonduras  =18442     
	msoLanguageIDMexicanSpanish   =2058      
	msoLanguageIDSpanishNicaragua =19466     
	msoLanguageIDSpanishPanama    =6154      
	msoLanguageIDSpanishParaguay  =15370     
	msoLanguageIDSpanishPeru      =10250     
	msoLanguageIDSpanishPuertoRico=20490     
	msoLanguageIDSpanishModernSort=3082      
	msoLanguageIDSpanish          =1034      
	msoLanguageIDSpanishUruguay   =14346     
	msoLanguageIDSpanishVenezuela =8202      
	msoLanguageIDSutu             =1072      
	msoLanguageIDSwahili          =1089      
	msoLanguageIDSwedishFinland   =2077      
	msoLanguageIDSwedish          =1053      
	msoLanguageIDSyriac           =1114      
	msoLanguageIDTajik            =1064      
	msoLanguageIDTamil            =1097      
	msoLanguageIDTamazight        =1119      
	msoLanguageIDTamazightLatin   =2143      
	msoLanguageIDTatar            =1092      
	msoLanguageIDTelugu           =1098      
	msoLanguageIDThai             =1054      
	msoLanguageIDTibetan          =1105      
	msoLanguageIDTigrignaEthiopic =1139      
	msoLanguageIDTigrignaEritrea  =2163      
	msoLanguageIDTsonga           =1073      
	msoLanguageIDTswana           =1074      
	msoLanguageIDTurkish          =1055      
	msoLanguageIDTurkmen          =1090      
	msoLanguageIDUkrainian        =1058      
	msoLanguageIDUrdu             =1056      
	msoLanguageIDUzbekCyrillic    =2115      
	msoLanguageIDUzbekLatin       =1091      
	msoLanguageIDVenda            =1075      
	msoLanguageIDVietnamese       =1066      
	msoLanguageIDWelsh            =1106      
	msoLanguageIDXhosa            =1076      
	msoLanguageIDYi               =1144      
	msoLanguageIDYiddish          =1085      
	msoLanguageIDYoruba           =1130      
	msoLanguageIDZulu             =1077      

class MsoLanguageIDHidden(Enum):
	msoLanguageIDChineseHongKong  =3076      
	msoLanguageIDChineseMacao     =5124      
	msoLanguageIDEnglishTrinidad  =11273     

class MsoLastModified(Enum):
	msoLastModifiedYesterday      =1         
	msoLastModifiedToday          =2         
	msoLastModifiedLastWeek       =3         
	msoLastModifiedThisWeek       =4         
	msoLastModifiedLastMonth      =5         
	msoLastModifiedThisMonth      =6         
	msoLastModifiedAnyTime        =7         

class MsoLightRigType(Enum):
	msoLightRigMixed              =-2        
	msoLightRigLegacyFlat1        =1         
	msoLightRigLegacyFlat2        =2         
	msoLightRigLegacyFlat3        =3         
	msoLightRigLegacyFlat4        =4         
	msoLightRigLegacyNormal1      =5         
	msoLightRigLegacyNormal2      =6         
	msoLightRigLegacyNormal3      =7         
	msoLightRigLegacyNormal4      =8         
	msoLightRigLegacyHarsh1       =9         
	msoLightRigLegacyHarsh2       =10        
	msoLightRigLegacyHarsh3       =11        
	msoLightRigLegacyHarsh4       =12        
	msoLightRigThreePoint         =13        
	msoLightRigBalanced           =14        
	msoLightRigSoft               =15        
	msoLightRigHarsh              =16        
	msoLightRigFlood              =17        
	msoLightRigContrasting        =18        
	msoLightRigMorning            =19        
	msoLightRigSunrise            =20        
	msoLightRigSunset             =21        
	msoLightRigChilly             =22        
	msoLightRigFreezing           =23        
	msoLightRigFlat               =24        
	msoLightRigTwoPoint           =25        
	msoLightRigGlow               =26        
	msoLightRigBrightRoom         =27        

class MsoLineCapStyle(Enum):
	msoLineCapMixed               =-2        
	msoLineCapSquare              =1         
	msoLineCapRound               =2         
	msoLineCapFlat                =3         

class MsoLineDashStyle(Enum):
	msoLineDashStyleMixed         =-2        
	msoLineSolid                  =1         
	msoLineSquareDot              =2         
	msoLineRoundDot               =3         
	msoLineDash                   =4         
	msoLineDashDot                =5         
	msoLineDashDotDot             =6         
	msoLineLongDash               =7         
	msoLineLongDashDot            =8         
	msoLineLongDashDotDot         =9         
	msoLineSysDash                =10        
	msoLineSysDot                 =11        
	msoLineSysDashDot             =12        

class MsoLineFillType(Enum):
	msoLineFillMixed              =-2        
	msoLineFillNone               =0         
	msoLineFillSolid              =1         
	msoLineFillPatterned          =2         
	msoLineFillGradient           =3         
	msoLineFillTextured           =4         
	msoLineFillBackground         =5         
	msoLineFillPicture            =6         

class MsoLineJoinStyle(Enum):
	msoLineJoinMixed              =-2        
	msoLineJoinRound              =1         
	msoLineJoinBevel              =2         
	msoLineJoinMiter              =3         

class MsoLineStyle(Enum):
	msoLineStyleMixed             =-2        
	msoLineSingle                 =1         
	msoLineThinThin               =2         
	msoLineThinThick              =3         
	msoLineThickThin              =4         
	msoLineThickBetweenThin       =5         

class MsoMenuAnimation(Enum):
	msoMenuAnimationNone          =0         
	msoMenuAnimationRandom        =1         
	msoMenuAnimationUnfold        =2         
	msoMenuAnimationSlide         =3         

class MsoMergeCmd(Enum):
	msoMergeUnion                 =1         
	msoMergeCombine               =2         
	msoMergeIntersect             =3         
	msoMergeSubtract              =4         
	msoMergeFragment              =5         

class MsoMetaPropertyType(Enum):
	msoMetaPropertyTypeUnknown    =0         
	msoMetaPropertyTypeBoolean    =1         
	msoMetaPropertyTypeChoice     =2         
	msoMetaPropertyTypeCalculated =3         
	msoMetaPropertyTypeComputed   =4         
	msoMetaPropertyTypeCurrency   =5         
	msoMetaPropertyTypeDateTime   =6         
	msoMetaPropertyTypeFillInChoice=7         
	msoMetaPropertyTypeGuid       =8         
	msoMetaPropertyTypeInteger    =9         
	msoMetaPropertyTypeLookup     =10        
	msoMetaPropertyTypeMultiChoiceLookup=11        
	msoMetaPropertyTypeMultiChoice=12        
	msoMetaPropertyTypeMultiChoiceFillIn=13        
	msoMetaPropertyTypeNote       =14        
	msoMetaPropertyTypeNumber     =15        
	msoMetaPropertyTypeText       =16        
	msoMetaPropertyTypeUrl        =17        
	msoMetaPropertyTypeUser       =18        
	msoMetaPropertyTypeUserMulti  =19        
	msoMetaPropertyTypeBusinessData=20        
	msoMetaPropertyTypeBusinessDataSecondary=21        
	msoMetaPropertyTypeMax        =22        

class MsoMixedType(Enum):
	msoIntegerMixed               =32768     
	msoSingleMixed                =-2147483648

class MsoModeType(Enum):
	msoModeModal                  =0         
	msoModeAutoDown               =1         
	msoModeModeless               =2         

class MsoMoveRow(Enum):
	msoMoveRowFirst               =-4        
	msoMoveRowPrev                =-3        
	msoMoveRowNext                =-2        
	msoMoveRowNbr                 =-1        

class MsoNumberedBulletStyle(Enum):
	msoBulletStyleMixed           =-2        
	msoBulletAlphaLCPeriod        =0         
	msoBulletAlphaUCPeriod        =1         
	msoBulletArabicParenRight     =2         
	msoBulletArabicPeriod         =3         
	msoBulletRomanLCParenBoth     =4         
	msoBulletRomanLCParenRight    =5         
	msoBulletRomanLCPeriod        =6         
	msoBulletRomanUCPeriod        =7         
	msoBulletAlphaLCParenBoth     =8         
	msoBulletAlphaLCParenRight    =9         
	msoBulletAlphaUCParenBoth     =10        
	msoBulletAlphaUCParenRight    =11        
	msoBulletArabicParenBoth      =12        
	msoBulletArabicPlain          =13        
	msoBulletRomanUCParenBoth     =14        
	msoBulletRomanUCParenRight    =15        
	msoBulletSimpChinPlain        =16        
	msoBulletSimpChinPeriod       =17        
	msoBulletCircleNumDBPlain     =18        
	msoBulletCircleNumWDWhitePlain=19        
	msoBulletCircleNumWDBlackPlain=20        
	msoBulletTradChinPlain        =21        
	msoBulletTradChinPeriod       =22        
	msoBulletArabicAlphaDash      =23        
	msoBulletArabicAbjadDash      =24        
	msoBulletHebrewAlphaDash      =25        
	msoBulletKanjiKoreanPlain     =26        
	msoBulletKanjiKoreanPeriod    =27        
	msoBulletArabicDBPlain        =28        
	msoBulletArabicDBPeriod       =29        
	msoBulletThaiAlphaPeriod      =30        
	msoBulletThaiAlphaParenRight  =31        
	msoBulletThaiAlphaParenBoth   =32        
	msoBulletThaiNumPeriod        =33        
	msoBulletThaiNumParenRight    =34        
	msoBulletThaiNumParenBoth     =35        
	msoBulletHindiAlphaPeriod     =36        
	msoBulletHindiNumPeriod       =37        
	msoBulletKanjiSimpChinDBPeriod=38        
	msoBulletHindiNumParenRight   =39        
	msoBulletHindiAlpha1Period    =40        

class MsoOLEMenuGroup(Enum):
	msoOLEMenuGroupNone           =-1        
	msoOLEMenuGroupFile           =0         
	msoOLEMenuGroupEdit           =1         
	msoOLEMenuGroupContainer      =2         
	msoOLEMenuGroupObject         =3         
	msoOLEMenuGroupWindow         =4         
	msoOLEMenuGroupHelp           =5         

class MsoOrgChartLayoutType(Enum):
	msoOrgChartLayoutMixed        =-2        
	msoOrgChartLayoutStandard     =1         
	msoOrgChartLayoutBothHanging  =2         
	msoOrgChartLayoutLeftHanging  =3         
	msoOrgChartLayoutRightHanging =4         
	msoOrgChartLayoutDefault      =5         

class MsoOrgChartOrientation(Enum):
	msoOrgChartOrientationMixed   =-2        
	msoOrgChartOrientationVertical=1         

class MsoOrientation(Enum):
	msoOrientationMixed           =-2        
	msoOrientationHorizontal      =1         
	msoOrientationVertical        =2         

class MsoParagraphAlignment(Enum):
	msoAlignMixed                 =-2        
	msoAlignLeft                  =1         
	msoAlignCenter                =2         
	msoAlignRight                 =3         
	msoAlignJustify               =4         
	msoAlignDistribute            =5         
	msoAlignThaiDistribute        =6         
	msoAlignJustifyLow            =7         

class MsoPathFormat(Enum):
	msoPathTypeMixed              =-2        
	msoPathTypeNone               =0         
	msoPathType1                  =1         
	msoPathType2                  =2         
	msoPathType3                  =3         
	msoPathType4                  =4         

class MsoPatternType(Enum):
	msoPatternMixed               =-2        
	msoPattern5Percent            =1         
	msoPattern10Percent           =2         
	msoPattern20Percent           =3         
	msoPattern25Percent           =4         
	msoPattern30Percent           =5         
	msoPattern40Percent           =6         
	msoPattern50Percent           =7         
	msoPattern60Percent           =8         
	msoPattern70Percent           =9         
	msoPattern75Percent           =10        
	msoPattern80Percent           =11        
	msoPattern90Percent           =12        
	msoPatternDarkHorizontal      =13        
	msoPatternDarkVertical        =14        
	msoPatternDarkDownwardDiagonal=15        
	msoPatternDarkUpwardDiagonal  =16        
	msoPatternSmallCheckerBoard   =17        
	msoPatternTrellis             =18        
	msoPatternLightHorizontal     =19        
	msoPatternLightVertical       =20        
	msoPatternLightDownwardDiagonal=21        
	msoPatternLightUpwardDiagonal =22        
	msoPatternSmallGrid           =23        
	msoPatternDottedDiamond       =24        
	msoPatternWideDownwardDiagonal=25        
	msoPatternWideUpwardDiagonal  =26        
	msoPatternDashedUpwardDiagonal=27        
	msoPatternDashedDownwardDiagonal=28        
	msoPatternNarrowVertical      =29        
	msoPatternNarrowHorizontal    =30        
	msoPatternDashedVertical      =31        
	msoPatternDashedHorizontal    =32        
	msoPatternLargeConfetti       =33        
	msoPatternLargeGrid           =34        
	msoPatternHorizontalBrick     =35        
	msoPatternLargeCheckerBoard   =36        
	msoPatternSmallConfetti       =37        
	msoPatternZigZag              =38        
	msoPatternSolidDiamond        =39        
	msoPatternDiagonalBrick       =40        
	msoPatternOutlinedDiamond     =41        
	msoPatternPlaid               =42        
	msoPatternSphere              =43        
	msoPatternWeave               =44        
	msoPatternDottedGrid          =45        
	msoPatternDivot               =46        
	msoPatternShingle             =47        
	msoPatternWave                =48        
	msoPatternHorizontal          =49        
	msoPatternVertical            =50        
	msoPatternCross               =51        
	msoPatternDownwardDiagonal    =52        
	msoPatternUpwardDiagonal      =53        
	msoPatternDiagonalCross       =54        

class MsoPermission(Enum):
	msoPermissionView             =1         
	msoPermissionRead             =1         
	msoPermissionEdit             =2         
	msoPermissionSave             =4         
	msoPermissionExtract          =8         
	msoPermissionChange           =15        
	msoPermissionPrint            =16        
	msoPermissionObjModel         =32        
	msoPermissionFullControl      =64        
	msoPermissionAllCommon        =127       

class MsoPickerField(Enum):
	msoPickerFieldUnknown         =0         
	msoPickerFieldDateTime        =1         
	msoPickerFieldNumber          =2         
	msoPickerFieldText            =3         
	msoPickerFieldUser            =4         
	msoPickerFieldMax             =5         

class MsoPictureColorType(Enum):
	msoPictureMixed               =-2        
	msoPictureAutomatic           =1         
	msoPictureGrayscale           =2         
	msoPictureBlackAndWhite       =3         
	msoPictureWatermark           =4         

class MsoPictureCompress(Enum):
	msoPictureCompressDocDefault  =-1        
	msoPictureCompressFalse       =0         
	msoPictureCompressTrue        =1         

class MsoPictureEffectType(Enum):
	msoEffectNone                 =0         
	msoEffectBackgroundRemoval    =1         
	msoEffectBlur                 =2         
	msoEffectBrightnessContrast   =3         
	msoEffectCement               =4         
	msoEffectCrisscrossEtching    =5         
	msoEffectChalkSketch          =6         
	msoEffectColorTemperature     =7         
	msoEffectCutout               =8         
	msoEffectFilmGrain            =9         
	msoEffectGlass                =10        
	msoEffectGlowDiffused         =11        
	msoEffectGlowEdges            =12        
	msoEffectLightScreen          =13        
	msoEffectLineDrawing          =14        
	msoEffectMarker               =15        
	msoEffectMosiaicBubbles       =16        
	msoEffectPaintBrush           =17        
	msoEffectPaintStrokes         =18        
	msoEffectPastelsSmooth        =19        
	msoEffectPencilGrayscale      =20        
	msoEffectPencilSketch         =21        
	msoEffectPhotocopy            =22        
	msoEffectPlasticWrap          =23        
	msoEffectSaturation           =24        
	msoEffectSharpenSoften        =25        
	msoEffectTexturizer           =26        
	msoEffectWatercolorSponge     =27        

class MsoPictureType(Enum):
	msoPictureTypeDefault         =-2        
	msoPictureTypePNG             =0         
	msoPictureTypeBMP             =1         
	msoPictureTypeGIF             =2         
	msoPictureTypeJPG             =3         
	msoPictureTypePDF             =4         

class MsoPresetCamera(Enum):
	msoPresetCameraMixed          =-2        
	msoCameraLegacyObliqueTopLeft =1         
	msoCameraLegacyObliqueTop     =2         
	msoCameraLegacyObliqueTopRight=3         
	msoCameraLegacyObliqueLeft    =4         
	msoCameraLegacyObliqueFront   =5         
	msoCameraLegacyObliqueRight   =6         
	msoCameraLegacyObliqueBottomLeft=7         
	msoCameraLegacyObliqueBottom  =8         
	msoCameraLegacyObliqueBottomRight=9         
	msoCameraLegacyPerspectiveTopLeft=10        
	msoCameraLegacyPerspectiveTop =11        
	msoCameraLegacyPerspectiveTopRight=12        
	msoCameraLegacyPerspectiveLeft=13        
	msoCameraLegacyPerspectiveFront=14        
	msoCameraLegacyPerspectiveRight=15        
	msoCameraLegacyPerspectiveBottomLeft=16        
	msoCameraLegacyPerspectiveBottom=17        
	msoCameraLegacyPerspectiveBottomRight=18        
	msoCameraOrthographicFront    =19        
	msoCameraIsometricTopUp       =20        
	msoCameraIsometricTopDown     =21        
	msoCameraIsometricBottomUp    =22        
	msoCameraIsometricBottomDown  =23        
	msoCameraIsometricLeftUp      =24        
	msoCameraIsometricLeftDown    =25        
	msoCameraIsometricRightUp     =26        
	msoCameraIsometricRightDown   =27        
	msoCameraIsometricOffAxis1Left=28        
	msoCameraIsometricOffAxis1Right=29        
	msoCameraIsometricOffAxis1Top =30        
	msoCameraIsometricOffAxis2Left=31        
	msoCameraIsometricOffAxis2Right=32        
	msoCameraIsometricOffAxis2Top =33        
	msoCameraIsometricOffAxis3Left=34        
	msoCameraIsometricOffAxis3Right=35        
	msoCameraIsometricOffAxis3Bottom=36        
	msoCameraIsometricOffAxis4Left=37        
	msoCameraIsometricOffAxis4Right=38        
	msoCameraIsometricOffAxis4Bottom=39        
	msoCameraObliqueTopLeft       =40        
	msoCameraObliqueTop           =41        
	msoCameraObliqueTopRight      =42        
	msoCameraObliqueLeft          =43        
	msoCameraObliqueRight         =44        
	msoCameraObliqueBottomLeft    =45        
	msoCameraObliqueBottom        =46        
	msoCameraObliqueBottomRight   =47        
	msoCameraPerspectiveFront     =48        
	msoCameraPerspectiveLeft      =49        
	msoCameraPerspectiveRight     =50        
	msoCameraPerspectiveAbove     =51        
	msoCameraPerspectiveBelow     =52        
	msoCameraPerspectiveAboveLeftFacing=53        
	msoCameraPerspectiveAboveRightFacing=54        
	msoCameraPerspectiveContrastingLeftFacing=55        
	msoCameraPerspectiveContrastingRightFacing=56        
	msoCameraPerspectiveHeroicLeftFacing=57        
	msoCameraPerspectiveHeroicRightFacing=58        
	msoCameraPerspectiveHeroicExtremeLeftFacing=59        
	msoCameraPerspectiveHeroicExtremeRightFacing=60        
	msoCameraPerspectiveRelaxed   =61        
	msoCameraPerspectiveRelaxedModerately=62        

class MsoPresetExtrusionDirection(Enum):
	msoPresetExtrusionDirectionMixed=-2        
	msoExtrusionBottomRight       =1         
	msoExtrusionBottom            =2         
	msoExtrusionBottomLeft        =3         
	msoExtrusionRight             =4         
	msoExtrusionNone              =5         
	msoExtrusionLeft              =6         
	msoExtrusionTopRight          =7         
	msoExtrusionTop               =8         
	msoExtrusionTopLeft           =9         

class MsoPresetGradientType(Enum):
	msoPresetGradientMixed        =-2        
	msoGradientEarlySunset        =1         
	msoGradientLateSunset         =2         
	msoGradientNightfall          =3         
	msoGradientDaybreak           =4         
	msoGradientHorizon            =5         
	msoGradientDesert             =6         
	msoGradientOcean              =7         
	msoGradientCalmWater          =8         
	msoGradientFire               =9         
	msoGradientFog                =10        
	msoGradientMoss               =11        
	msoGradientPeacock            =12        
	msoGradientWheat              =13        
	msoGradientParchment          =14        
	msoGradientMahogany           =15        
	msoGradientRainbow            =16        
	msoGradientRainbowII          =17        
	msoGradientGold               =18        
	msoGradientGoldII             =19        
	msoGradientBrass              =20        
	msoGradientChrome             =21        
	msoGradientChromeII           =22        
	msoGradientSilver             =23        
	msoGradientSapphire           =24        

class MsoPresetLightingDirection(Enum):
	msoPresetLightingDirectionMixed=-2        
	msoLightingTopLeft            =1         
	msoLightingTop                =2         
	msoLightingTopRight           =3         
	msoLightingLeft               =4         
	msoLightingNone               =5         
	msoLightingRight              =6         
	msoLightingBottomLeft         =7         
	msoLightingBottom             =8         
	msoLightingBottomRight        =9         

class MsoPresetLightingSoftness(Enum):
	msoPresetLightingSoftnessMixed=-2        
	msoLightingDim                =1         
	msoLightingNormal             =2         
	msoLightingBright             =3         

class MsoPresetMaterial(Enum):
	msoPresetMaterialMixed        =-2        
	msoMaterialMatte              =1         
	msoMaterialPlastic            =2         
	msoMaterialMetal              =3         
	msoMaterialWireFrame          =4         
	msoMaterialMatte2             =5         
	msoMaterialPlastic2           =6         
	msoMaterialMetal2             =7         
	msoMaterialWarmMatte          =8         
	msoMaterialTranslucentPowder  =9         
	msoMaterialPowder             =10        
	msoMaterialDarkEdge           =11        
	msoMaterialSoftEdge           =12        
	msoMaterialClear              =13        
	msoMaterialFlat               =14        
	msoMaterialSoftMetal          =15        

class MsoPresetTextEffect(Enum):
	msoTextEffectMixed            =-2        
	msoTextEffect1                =0         
	msoTextEffect2                =1         
	msoTextEffect3                =2         
	msoTextEffect4                =3         
	msoTextEffect5                =4         
	msoTextEffect6                =5         
	msoTextEffect7                =6         
	msoTextEffect8                =7         
	msoTextEffect9                =8         
	msoTextEffect10               =9         
	msoTextEffect11               =10        
	msoTextEffect12               =11        
	msoTextEffect13               =12        
	msoTextEffect14               =13        
	msoTextEffect15               =14        
	msoTextEffect16               =15        
	msoTextEffect17               =16        
	msoTextEffect18               =17        
	msoTextEffect19               =18        
	msoTextEffect20               =19        
	msoTextEffect21               =20        
	msoTextEffect22               =21        
	msoTextEffect23               =22        
	msoTextEffect24               =23        
	msoTextEffect25               =24        
	msoTextEffect26               =25        
	msoTextEffect27               =26        
	msoTextEffect28               =27        
	msoTextEffect29               =28        
	msoTextEffect30               =29        
	msoTextEffect31               =30        
	msoTextEffect32               =31        
	msoTextEffect33               =32        
	msoTextEffect34               =33        
	msoTextEffect35               =34        
	msoTextEffect36               =35        
	msoTextEffect37               =36        
	msoTextEffect38               =37        
	msoTextEffect39               =38        
	msoTextEffect40               =39        
	msoTextEffect41               =40        
	msoTextEffect42               =41        
	msoTextEffect43               =42        
	msoTextEffect44               =43        
	msoTextEffect45               =44        
	msoTextEffect46               =45        
	msoTextEffect47               =46        
	msoTextEffect48               =47        
	msoTextEffect49               =48        
	msoTextEffect50               =49        

class MsoPresetTextEffectShape(Enum):
	msoTextEffectShapeMixed       =-2        
	msoTextEffectShapePlainText   =1         
	msoTextEffectShapeStop        =2         
	msoTextEffectShapeTriangleUp  =3         
	msoTextEffectShapeTriangleDown=4         
	msoTextEffectShapeChevronUp   =5         
	msoTextEffectShapeChevronDown =6         
	msoTextEffectShapeRingInside  =7         
	msoTextEffectShapeRingOutside =8         
	msoTextEffectShapeArchUpCurve =9         
	msoTextEffectShapeArchDownCurve=10        
	msoTextEffectShapeCircleCurve =11        
	msoTextEffectShapeButtonCurve =12        
	msoTextEffectShapeArchUpPour  =13        
	msoTextEffectShapeArchDownPour=14        
	msoTextEffectShapeCirclePour  =15        
	msoTextEffectShapeButtonPour  =16        
	msoTextEffectShapeCurveUp     =17        
	msoTextEffectShapeCurveDown   =18        
	msoTextEffectShapeCanUp       =19        
	msoTextEffectShapeCanDown     =20        
	msoTextEffectShapeWave1       =21        
	msoTextEffectShapeWave2       =22        
	msoTextEffectShapeDoubleWave1 =23        
	msoTextEffectShapeDoubleWave2 =24        
	msoTextEffectShapeInflate     =25        
	msoTextEffectShapeDeflate     =26        
	msoTextEffectShapeInflateBottom=27        
	msoTextEffectShapeDeflateBottom=28        
	msoTextEffectShapeInflateTop  =29        
	msoTextEffectShapeDeflateTop  =30        
	msoTextEffectShapeDeflateInflate=31        
	msoTextEffectShapeDeflateInflateDeflate=32        
	msoTextEffectShapeFadeRight   =33        
	msoTextEffectShapeFadeLeft    =34        
	msoTextEffectShapeFadeUp      =35        
	msoTextEffectShapeFadeDown    =36        
	msoTextEffectShapeSlantUp     =37        
	msoTextEffectShapeSlantDown   =38        
	msoTextEffectShapeCascadeUp   =39        
	msoTextEffectShapeCascadeDown =40        

class MsoPresetTexture(Enum):
	msoPresetTextureMixed         =-2        
	msoTexturePapyrus             =1         
	msoTextureCanvas              =2         
	msoTextureDenim               =3         
	msoTextureWovenMat            =4         
	msoTextureWaterDroplets       =5         
	msoTexturePaperBag            =6         
	msoTextureFishFossil          =7         
	msoTextureSand                =8         
	msoTextureGreenMarble         =9         
	msoTextureWhiteMarble         =10        
	msoTextureBrownMarble         =11        
	msoTextureGranite             =12        
	msoTextureNewsprint           =13        
	msoTextureRecycledPaper       =14        
	msoTextureParchment           =15        
	msoTextureStationery          =16        
	msoTextureBlueTissuePaper     =17        
	msoTexturePinkTissuePaper     =18        
	msoTexturePurpleMesh          =19        
	msoTextureBouquet             =20        
	msoTextureCork                =21        
	msoTextureWalnut              =22        
	msoTextureOak                 =23        
	msoTextureMediumWood          =24        

class MsoPresetThreeDFormat(Enum):
	msoPresetThreeDFormatMixed    =-2        
	msoThreeD1                    =1         
	msoThreeD2                    =2         
	msoThreeD3                    =3         
	msoThreeD4                    =4         
	msoThreeD5                    =5         
	msoThreeD6                    =6         
	msoThreeD7                    =7         
	msoThreeD8                    =8         
	msoThreeD9                    =9         
	msoThreeD10                   =10        
	msoThreeD11                   =11        
	msoThreeD12                   =12        
	msoThreeD13                   =13        
	msoThreeD14                   =14        
	msoThreeD15                   =15        
	msoThreeD16                   =16        
	msoThreeD17                   =17        
	msoThreeD18                   =18        
	msoThreeD19                   =19        
	msoThreeD20                   =20        

class MsoPrivacyOptionSetting(Enum):
	msoUserContentEnabled         =0         
	msoDownloadContentEnabled     =1         
	msoControllerConnectedServicesEnabled=2         
	msoDisconnectedState          =3         

class MsoReflectionType(Enum):
	msoReflectionTypeMixed        =-2        
	msoReflectionTypeNone         =0         
	msoReflectionType1            =1         
	msoReflectionType2            =2         
	msoReflectionType3            =3         
	msoReflectionType4            =4         
	msoReflectionType5            =5         
	msoReflectionType6            =6         
	msoReflectionType7            =7         
	msoReflectionType8            =8         
	msoReflectionType9            =9         

class MsoRelativeNodePosition(Enum):
	msoBeforeNode                 =1         
	msoAfterNode                  =2         
	msoBeforeFirstSibling         =3         
	msoAfterLastSibling           =4         

class MsoScaleFrom(Enum):
	msoScaleFromTopLeft           =0         
	msoScaleFromMiddle            =1         
	msoScaleFromBottomRight       =2         

class MsoScreenSize(Enum):
	msoScreenSize544x376          =0         
	msoScreenSize640x480          =1         
	msoScreenSize720x512          =2         
	msoScreenSize800x600          =3         
	msoScreenSize1024x768         =4         
	msoScreenSize1152x882         =5         
	msoScreenSize1152x900         =6         
	msoScreenSize1280x1024        =7         
	msoScreenSize1600x1200        =8         
	msoScreenSize1800x1440        =9         
	msoScreenSize1920x1200        =10        

class MsoScriptLanguage(Enum):
	msoScriptLanguageJava         =1         
	msoScriptLanguageVisualBasic  =2         
	msoScriptLanguageASP          =3         
	msoScriptLanguageOther        =4         

class MsoScriptLocation(Enum):
	msoScriptLocationInHead       =1         
	msoScriptLocationInBody       =2         

class MsoSearchIn(Enum):
	msoSearchInMyComputer         =0         
	msoSearchInOutlook            =1         
	msoSearchInMyNetworkPlaces    =2         
	msoSearchInCustom             =3         

class MsoSegmentType(Enum):
	msoSegmentLine                =0         
	msoSegmentCurve               =1         

class MsoSensitivityLabelError(Enum):
	msoNoError                    =0         
	msoUserNotSignedIn            =1         

class MsoShadowStyle(Enum):
	msoShadowStyleMixed           =-2        
	msoShadowStyleInnerShadow     =1         
	msoShadowStyleOuterShadow     =2         

class MsoShadowType(Enum):
	msoShadowMixed                =-2        
	msoShadow1                    =1         
	msoShadow2                    =2         
	msoShadow3                    =3         
	msoShadow4                    =4         
	msoShadow5                    =5         
	msoShadow6                    =6         
	msoShadow7                    =7         
	msoShadow8                    =8         
	msoShadow9                    =9         
	msoShadow10                   =10        
	msoShadow11                   =11        
	msoShadow12                   =12        
	msoShadow13                   =13        
	msoShadow14                   =14        
	msoShadow15                   =15        
	msoShadow16                   =16        
	msoShadow17                   =17        
	msoShadow18                   =18        
	msoShadow19                   =19        
	msoShadow20                   =20        
	msoShadow21                   =21        
	msoShadow22                   =22        
	msoShadow23                   =23        
	msoShadow24                   =24        
	msoShadow25                   =25        
	msoShadow26                   =26        
	msoShadow27                   =27        
	msoShadow28                   =28        
	msoShadow29                   =29        
	msoShadow30                   =30        
	msoShadow31                   =31        
	msoShadow32                   =32        
	msoShadow33                   =33        
	msoShadow34                   =34        
	msoShadow35                   =35        
	msoShadow36                   =36        
	msoShadow37                   =37        
	msoShadow38                   =38        
	msoShadow39                   =39        
	msoShadow40                   =40        
	msoShadow41                   =41        
	msoShadow42                   =42        
	msoShadow43                   =43        

class MsoShapeStyleIndex(Enum):
	msoShapeStyleMixed            =-2        
	msoShapeStyleNotAPreset       =0         
	msoShapeStylePreset1          =1         
	msoShapeStylePreset2          =2         
	msoShapeStylePreset3          =3         
	msoShapeStylePreset4          =4         
	msoShapeStylePreset5          =5         
	msoShapeStylePreset6          =6         
	msoShapeStylePreset7          =7         
	msoShapeStylePreset8          =8         
	msoShapeStylePreset9          =9         
	msoShapeStylePreset10         =10        
	msoShapeStylePreset11         =11        
	msoShapeStylePreset12         =12        
	msoShapeStylePreset13         =13        
	msoShapeStylePreset14         =14        
	msoShapeStylePreset15         =15        
	msoShapeStylePreset16         =16        
	msoShapeStylePreset17         =17        
	msoShapeStylePreset18         =18        
	msoShapeStylePreset19         =19        
	msoShapeStylePreset20         =20        
	msoShapeStylePreset21         =21        
	msoShapeStylePreset22         =22        
	msoShapeStylePreset23         =23        
	msoShapeStylePreset24         =24        
	msoShapeStylePreset25         =25        
	msoShapeStylePreset26         =26        
	msoShapeStylePreset27         =27        
	msoShapeStylePreset28         =28        
	msoShapeStylePreset29         =29        
	msoShapeStylePreset30         =30        
	msoShapeStylePreset31         =31        
	msoShapeStylePreset32         =32        
	msoShapeStylePreset33         =33        
	msoShapeStylePreset34         =34        
	msoShapeStylePreset35         =35        
	msoShapeStylePreset36         =36        
	msoShapeStylePreset37         =37        
	msoShapeStylePreset38         =38        
	msoShapeStylePreset39         =39        
	msoShapeStylePreset40         =40        
	msoShapeStylePreset41         =41        
	msoShapeStylePreset42         =42        
	msoShapeStylePreset43         =43        
	msoShapeStylePreset44         =44        
	msoShapeStylePreset45         =45        
	msoShapeStylePreset46         =46        
	msoShapeStylePreset47         =47        
	msoShapeStylePreset48         =48        
	msoShapeStylePreset49         =49        
	msoShapeStylePreset50         =50        
	msoShapeStylePreset51         =51        
	msoShapeStylePreset52         =52        
	msoShapeStylePreset53         =53        
	msoShapeStylePreset54         =54        
	msoShapeStylePreset55         =55        
	msoShapeStylePreset56         =56        
	msoShapeStylePreset57         =57        
	msoShapeStylePreset58         =58        
	msoShapeStylePreset59         =59        
	msoShapeStylePreset60         =60        
	msoShapeStylePreset61         =61        
	msoShapeStylePreset62         =62        
	msoShapeStylePreset63         =63        
	msoShapeStylePreset64         =64        
	msoShapeStylePreset65         =65        
	msoShapeStylePreset66         =66        
	msoShapeStylePreset67         =67        
	msoShapeStylePreset68         =68        
	msoShapeStylePreset69         =69        
	msoShapeStylePreset70         =70        
	msoShapeStylePreset71         =71        
	msoShapeStylePreset72         =72        
	msoShapeStylePreset73         =73        
	msoShapeStylePreset74         =74        
	msoShapeStylePreset75         =75        
	msoShapeStylePreset76         =76        
	msoShapeStylePreset77         =77        
	msoLineStylePreset1           =10001     
	msoLineStylePreset2           =10002     
	msoLineStylePreset3           =10003     
	msoLineStylePreset4           =10004     
	msoLineStylePreset5           =10005     
	msoLineStylePreset6           =10006     
	msoLineStylePreset7           =10007     
	msoLineStylePreset8           =10008     
	msoLineStylePreset9           =10009     
	msoLineStylePreset10          =10010     
	msoLineStylePreset11          =10011     
	msoLineStylePreset12          =10012     
	msoLineStylePreset13          =10013     
	msoLineStylePreset14          =10014     
	msoLineStylePreset15          =10015     
	msoLineStylePreset16          =10016     
	msoLineStylePreset17          =10017     
	msoLineStylePreset18          =10018     
	msoLineStylePreset19          =10019     
	msoLineStylePreset20          =10020     
	msoLineStylePreset21          =10021     
	msoLineStylePreset22          =10022     
	msoLineStylePreset23          =10023     
	msoLineStylePreset24          =10024     
	msoLineStylePreset25          =10025     
	msoLineStylePreset26          =10026     
	msoLineStylePreset27          =10027     
	msoLineStylePreset28          =10028     
	msoLineStylePreset29          =10029     
	msoLineStylePreset30          =10030     
	msoLineStylePreset31          =10031     
	msoLineStylePreset32          =10032     
	msoLineStylePreset33          =10033     
	msoLineStylePreset34          =10034     
	msoLineStylePreset35          =10035     
	msoLineStylePreset36          =10036     
	msoLineStylePreset37          =10037     
	msoLineStylePreset38          =10038     
	msoLineStylePreset39          =10039     
	msoLineStylePreset40          =10040     
	msoLineStylePreset41          =10041     
	msoLineStylePreset42          =10042     

class MsoShapeType(Enum):
	msoShapeTypeMixed             =-2        
	msoAutoShape                  =1         
	msoCallout                    =2         
	msoChart                      =3         
	msoComment                    =4         
	msoFreeform                   =5         
	msoGroup                      =6         
	msoEmbeddedOLEObject          =7         
	msoFormControl                =8         
	msoLine                       =9         
	msoLinkedOLEObject            =10        
	msoLinkedPicture              =11        
	msoOLEControlObject           =12        
	msoPicture                    =13        
	msoPlaceholder                =14        
	msoTextEffect                 =15        
	msoMedia                      =16        
	msoTextBox                    =17        
	msoScriptAnchor               =18        
	msoTable                      =19        
	msoCanvas                     =20        
	msoDiagram                    =21        
	msoInk                        =22        
	msoInkComment                 =23        
	msoSmartArt                   =24        
	msoSlicer                     =25        
	msoWebVideo                   =26        
	msoContentApp                 =27        
	msoGraphic                    =28        
	msoLinkedGraphic              =29        
	mso3DModel                    =30        
	msoLinked3DModel              =31        

class MsoSharedWorkspaceTaskPriority(Enum):
	msoSharedWorkspaceTaskPriorityHigh=1         
	msoSharedWorkspaceTaskPriorityNormal=2         
	msoSharedWorkspaceTaskPriorityLow=3         

class MsoSharedWorkspaceTaskStatus(Enum):
	msoSharedWorkspaceTaskStatusNotStarted=1         
	msoSharedWorkspaceTaskStatusInProgress=2         
	msoSharedWorkspaceTaskStatusCompleted=3         
	msoSharedWorkspaceTaskStatusDeferred=4         
	msoSharedWorkspaceTaskStatusWaiting=5         

class MsoSignatureSubset(Enum):
	msoSignatureSubsetSignaturesAllSigs=0         
	msoSignatureSubsetSignaturesNonVisible=1         
	msoSignatureSubsetSignatureLines=2         
	msoSignatureSubsetSignatureLinesSigned=3         
	msoSignatureSubsetSignatureLinesUnsigned=4         
	msoSignatureSubsetAll         =5         

class MsoSmartArtNodePosition(Enum):
	msoSmartArtNodeDefault        =1         
	msoSmartArtNodeAfter          =2         
	msoSmartArtNodeBefore         =3         
	msoSmartArtNodeAbove          =4         
	msoSmartArtNodeBelow          =5         

class MsoSmartArtNodeType(Enum):
	msoSmartArtNodeTypeDefault    =1         
	msoSmartArtNodeTypeAssistant  =2         

class MsoSoftEdgeType(Enum):
	msoSoftEdgeTypeMixed          =-2        
	msoSoftEdgeTypeNone           =0         
	msoSoftEdgeType1              =1         
	msoSoftEdgeType2              =2         
	msoSoftEdgeType3              =3         
	msoSoftEdgeType4              =4         
	msoSoftEdgeType5              =5         
	msoSoftEdgeType6              =6         

class MsoSortBy(Enum):
	msoSortByFileName             =1         
	msoSortBySize                 =2         
	msoSortByFileType             =3         
	msoSortByLastModified         =4         
	msoSortByNone                 =5         

class MsoSortOrder(Enum):
	msoSortOrderAscending         =1         
	msoSortOrderDescending        =2         

class MsoSyncAvailableType(Enum):
	msoSyncAvailableNone          =0         
	msoSyncAvailableOffline       =1         
	msoSyncAvailableAnywhere      =2         

class MsoSyncCompareType(Enum):
	msoSyncCompareAndMerge        =0         
	msoSyncCompareSideBySide      =1         

class MsoSyncConflictResolutionType(Enum):
	msoSyncConflictClientWins     =0         
	msoSyncConflictServerWins     =1         
	msoSyncConflictMerge          =2         

class MsoSyncErrorType(Enum):
	msoSyncErrorNone              =0         
	msoSyncErrorUnauthorizedUser  =1         
	msoSyncErrorCouldNotConnect   =2         
	msoSyncErrorOutOfSpace        =3         
	msoSyncErrorFileNotFound      =4         
	msoSyncErrorFileTooLarge      =5         
	msoSyncErrorFileInUse         =6         
	msoSyncErrorVirusUpload       =7         
	msoSyncErrorVirusDownload     =8         
	msoSyncErrorUnknownUpload     =9         
	msoSyncErrorUnknownDownload   =10        
	msoSyncErrorCouldNotOpen      =11        
	msoSyncErrorCouldNotUpdate    =12        
	msoSyncErrorCouldNotCompare   =13        
	msoSyncErrorCouldNotResolve   =14        
	msoSyncErrorNoNetwork         =15        
	msoSyncErrorUnknown           =16        

class MsoSyncEventType(Enum):
	msoSyncEventDownloadInitiated =0         
	msoSyncEventDownloadSucceeded =1         
	msoSyncEventDownloadFailed    =2         
	msoSyncEventUploadInitiated   =3         
	msoSyncEventUploadSucceeded   =4         
	msoSyncEventUploadFailed      =5         
	msoSyncEventDownloadNoChange  =6         
	msoSyncEventOffline           =7         

class MsoSyncStatusType(Enum):
	msoSyncStatusNoSharedWorkspace=0         
	msoSyncStatusNotRoaming       =0         
	msoSyncStatusLatest           =1         
	msoSyncStatusNewerAvailable   =2         
	msoSyncStatusLocalChanges     =3         
	msoSyncStatusConflict         =4         
	msoSyncStatusSuspended        =5         
	msoSyncStatusError            =6         

class MsoSyncVersionType(Enum):
	msoSyncVersionLastViewed      =0         
	msoSyncVersionServer          =1         

class MsoTabStopType(Enum):
	msoTabStopMixed               =-2        
	msoTabStopLeft                =1         
	msoTabStopCenter              =2         
	msoTabStopRight               =3         
	msoTabStopDecimal             =4         

class MsoTargetBrowser(Enum):
	msoTargetBrowserV3            =0         
	msoTargetBrowserV4            =1         
	msoTargetBrowserIE4           =2         
	msoTargetBrowserIE5           =3         
	msoTargetBrowserIE6           =4         

class MsoTelemetryConsentLevel(Enum):
	msoTelemetryConsentLevelUnset =0         
	msoTelemetryConsentLevelBasic =1         
	msoTelemetryConsentLevelFull  =2         
	msoTelemetryConsentLevelZero  =3         
	msoTelemetryConsentLevelDefault=4         

class MsoTextCaps(Enum):
	msoCapsMixed                  =-2        
	msoNoCaps                     =0         
	msoSmallCaps                  =1         
	msoAllCaps                    =2         

class MsoTextChangeCase(Enum):
	msoCaseSentence               =1         
	msoCaseLower                  =2         
	msoCaseUpper                  =3         
	msoCaseTitle                  =4         
	msoCaseToggle                 =5         

class MsoTextCharWrap(Enum):
	msoCharWrapMixed              =-2        
	msoNoCharWrap                 =0         
	msoStandardCharWrap           =1         
	msoStrictCharWrap             =2         
	msoCustomCharWrap             =3         

class MsoTextDirection(Enum):
	msoTextDirectionMixed         =-2        
	msoTextDirectionLeftToRight   =1         
	msoTextDirectionRightToLeft   =2         

class MsoTextEffectAlignment(Enum):
	msoTextEffectAlignmentMixed   =-2        
	msoTextEffectAlignmentLeft    =1         
	msoTextEffectAlignmentCentered=2         
	msoTextEffectAlignmentRight   =3         
	msoTextEffectAlignmentLetterJustify=4         
	msoTextEffectAlignmentWordJustify=5         
	msoTextEffectAlignmentStretchJustify=6         

class MsoTextFontAlign(Enum):
	msoFontAlignMixed             =-2        
	msoFontAlignAuto              =0         
	msoFontAlignTop               =1         
	msoFontAlignCenter            =2         
	msoFontAlignBaseline          =3         
	msoFontAlignBottom            =4         

class MsoTextOrientation(Enum):
	msoTextOrientationMixed       =-2        
	msoTextOrientationHorizontal  =1         
	msoTextOrientationUpward      =2         
	msoTextOrientationDownward    =3         
	msoTextOrientationVerticalFarEast=4         
	msoTextOrientationVertical    =5         
	msoTextOrientationHorizontalRotatedFarEast=6         

class MsoTextRangeInsertPosition(Enum):
	msoMsoTextRangeInsertBefore   =0         
	msoMsoTextRangeInsertAfter    =1         

class MsoTextStrike(Enum):
	msoStrikeMixed                =-2        
	msoNoStrike                   =0         
	msoSingleStrike               =1         
	msoDoubleStrike               =2         

class MsoTextTabAlign(Enum):
	msoTabAlignMixed              =-2        
	msoTabAlignLeft               =0         
	msoTabAlignCenter             =1         
	msoTabAlignRight              =2         
	msoTabAlignDecimal            =3         

class MsoTextUnderlineType(Enum):
	msoUnderlineMixed             =-2        
	msoNoUnderline                =0         
	msoUnderlineWords             =1         
	msoUnderlineSingleLine        =2         
	msoUnderlineDoubleLine        =3         
	msoUnderlineHeavyLine         =4         
	msoUnderlineDottedLine        =5         
	msoUnderlineDottedHeavyLine   =6         
	msoUnderlineDashLine          =7         
	msoUnderlineDashHeavyLine     =8         
	msoUnderlineDashLongLine      =9         
	msoUnderlineDashLongHeavyLine =10        
	msoUnderlineDotDashLine       =11        
	msoUnderlineDotDashHeavyLine  =12        
	msoUnderlineDotDotDashLine    =13        
	msoUnderlineDotDotDashHeavyLine=14        
	msoUnderlineWavyLine          =15        
	msoUnderlineWavyHeavyLine     =16        
	msoUnderlineWavyDoubleLine    =17        

class MsoTextureAlignment(Enum):
	msoTextureAlignmentMixed      =-2        
	msoTextureTopLeft             =0         
	msoTextureTop                 =1         
	msoTextureTopRight            =2         
	msoTextureLeft                =3         
	msoTextureCenter              =4         
	msoTextureRight               =5         
	msoTextureBottomLeft          =6         
	msoTextureBottom              =7         
	msoTextureBottomRight         =8         

class MsoTextureType(Enum):
	msoTextureTypeMixed           =-2        
	msoTexturePreset              =1         
	msoTextureUserDefined         =2         

class MsoThemeColorIndex(Enum):
	msoThemeColorMixed            =-2        
	msoNotThemeColor              =0         
	msoThemeColorDark1            =1         
	msoThemeColorLight1           =2         
	msoThemeColorDark2            =3         
	msoThemeColorLight2           =4         
	msoThemeColorAccent1          =5         
	msoThemeColorAccent2          =6         
	msoThemeColorAccent3          =7         
	msoThemeColorAccent4          =8         
	msoThemeColorAccent5          =9         
	msoThemeColorAccent6          =10        
	msoThemeColorHyperlink        =11        
	msoThemeColorFollowedHyperlink=12        
	msoThemeColorText1            =13        
	msoThemeColorBackground1      =14        
	msoThemeColorText2            =15        
	msoThemeColorBackground2      =16        

class MsoThemeColorSchemeIndex(Enum):
	msoThemeDark1                 =1         
	msoThemeLight1                =2         
	msoThemeDark2                 =3         
	msoThemeLight2                =4         
	msoThemeAccent1               =5         
	msoThemeAccent2               =6         
	msoThemeAccent3               =7         
	msoThemeAccent4               =8         
	msoThemeAccent5               =9         
	msoThemeAccent6               =10        
	msoThemeHyperlink             =11        
	msoThemeFollowedHyperlink     =12        

class MsoTriState(Enum):
	msoTrue                       =-1        
	msoFalse                      =0         
	msoCTrue                      =1         
	msoTriStateToggle             =-3        
	msoTriStateMixed              =-2        

class MsoVerticalAnchor(Enum):
	msoVerticalAnchorMixed        =-2        
	msoAnchorTop                  =1         
	msoAnchorTopBaseline          =2         
	msoAnchorMiddle               =3         
	msoAnchorBottom               =4         
	msoAnchorBottomBaseLine       =5         

class MsoWarpFormat(Enum):
	msoWarpFormatMixed            =-2        
	msoWarpFormat1                =0         
	msoWarpFormat2                =1         
	msoWarpFormat3                =2         
	msoWarpFormat4                =3         
	msoWarpFormat5                =4         
	msoWarpFormat6                =5         
	msoWarpFormat7                =6         
	msoWarpFormat8                =7         
	msoWarpFormat9                =8         
	msoWarpFormat10               =9         
	msoWarpFormat11               =10        
	msoWarpFormat12               =11        
	msoWarpFormat13               =12        
	msoWarpFormat14               =13        
	msoWarpFormat15               =14        
	msoWarpFormat16               =15        
	msoWarpFormat17               =16        
	msoWarpFormat18               =17        
	msoWarpFormat19               =18        
	msoWarpFormat20               =19        
	msoWarpFormat21               =20        
	msoWarpFormat22               =21        
	msoWarpFormat23               =22        
	msoWarpFormat24               =23        
	msoWarpFormat25               =24        
	msoWarpFormat26               =25        
	msoWarpFormat27               =26        
	msoWarpFormat28               =27        
	msoWarpFormat29               =28        
	msoWarpFormat30               =29        
	msoWarpFormat31               =30        
	msoWarpFormat32               =31        
	msoWarpFormat33               =32        
	msoWarpFormat34               =33        
	msoWarpFormat35               =34        
	msoWarpFormat36               =35        
	msoWarpFormat37               =36        

class MsoWizardActType(Enum):
	msoWizardActInactive          =0         
	msoWizardActActive            =1         
	msoWizardActSuspend           =2         
	msoWizardActResume            =3         

class MsoWizardMsgType(Enum):
	msoWizardMsgLocalStateOn      =1         
	msoWizardMsgLocalStateOff     =2         
	msoWizardMsgShowHelp          =3         
	msoWizardMsgSuspending        =4         
	msoWizardMsgResuming          =5         

class MsoZOrderCmd(Enum):
	msoBringToFront               =0         
	msoSendToBack                 =1         
	msoBringForward               =2         
	msoSendBackward               =3         
	msoBringInFrontOfText         =4         
	msoSendBehindText             =5         

class RibbonControlSize(Enum):
	RibbonControlSizeRegular      =0         
	RibbonControlSizeLarge        =1         

class SignatureDetail(Enum):
	sigdetLocalSigningTime        =0         
	sigdetApplicationName         =1         
	sigdetApplicationVersion      =2         
	sigdetOfficeVersion           =3         
	sigdetWindowsVersion          =4         
	sigdetNumberOfMonitors        =5         
	sigdetHorizResolution         =6         
	sigdetVertResolution          =7         
	sigdetColorDepth              =8         
	sigdetSignedData              =9         
	sigdetDocPreviewImg           =10        
	sigdetIPFormHash              =11        
	sigdetIPCurrentView           =12        
	sigdetSignatureType           =13        
	sigdetHashAlgorithm           =14        
	sigdetShouldShowViewWarning   =15        
	sigdetDelSuggSigner           =16        
	sigdetDelSuggSignerSet        =17        
	sigdetDelSuggSignerLine2      =18        
	sigdetDelSuggSignerLine2Set   =19        
	sigdetDelSuggSignerEmail      =20        
	sigdetDelSuggSignerEmailSet   =21        

class SignatureLineImage(Enum):
	siglnimgSoftwareRequired      =0         
	siglnimgUnsigned              =1         
	siglnimgSignedValid           =2         
	siglnimgSignedInvalid         =3         
	siglnimgSigned                =4         

class SignatureProviderDetail(Enum):
	sigprovdetUrl                 =0         
	sigprovdetHashAlgorithm       =1         
	sigprovdetUIOnly              =2         
	sigprovdetUseOfficeUI         =3         
	sigprovdetUseOfficeStampUI    =4         

class SignatureType(Enum):
	sigtypeUnknown                =0         
	sigtypeNonVisible             =1         
	sigtypeSignatureLine          =2         
	sigtypeMax                    =3         

class XlAxisCrosses(Enum):
	xlAxisCrossesAutomatic        =-4105     
	xlAxisCrossesCustom           =-4114     
	xlAxisCrossesMaximum          =2         
	xlAxisCrossesMinimum          =4         

class XlAxisGroup(Enum):
	xlPrimary                     =1         
	xlSecondary                   =2         

class XlAxisType(Enum):
	xlCategory                    =1         
	xlSeriesAxis                  =3         
	xlValue                       =2         

class XlBarShape(Enum):
	xlBox                         =0         
	xlPyramidToPoint              =1         
	xlPyramidToMax                =2         
	xlCylinder                    =3         
	xlConeToPoint                 =4         
	xlConeToMax                   =5         

class XlBinsType(Enum):
	xlBinsTypeAutomatic           =0         
	xlBinsTypeCategorical         =1         
	xlBinsTypeManual              =2         
	xlBinsTypeBinSize             =3         
	xlBinsTypeBinCount            =4         

class XlBorderWeight(Enum):
	xlHairline                    =1         
	xlMedium                      =-4138     
	xlThick                       =4         
	xlThin                        =2         

class XlCategoryLabelLevel(Enum):
	xlCategoryLabelLevelNone      =-3        
	xlCategoryLabelLevelCustom    =-2        
	xlCategoryLabelLevelAll       =-1        

class XlCategorySortOrder(Enum):
	xlIndexAscending              =0         
	xlIndexDescending             =1         
	xlCategoryAscending           =2         
	xlCategoryDescending          =3         

class XlCategoryType(Enum):
	xlCategoryScale               =2         
	xlTimeScale                   =3         
	xlAutomaticScale              =-4105     

class XlChartElementPosition(Enum):
	xlChartElementPositionAutomatic=-4105     
	xlChartElementPositionCustom  =-4114     

class XlChartItem(Enum):
	xlDataLabel                   =0         
	xlChartArea                   =2         
	xlSeries                      =3         
	xlChartTitle                  =4         
	xlWalls                       =5         
	xlCorners                     =6         
	xlDataTable                   =7         
	xlTrendline                   =8         
	xlErrorBars                   =9         
	xlXErrorBars                  =10        
	xlYErrorBars                  =11        
	xlLegendEntry                 =12        
	xlLegendKey                   =13        
	xlShape                       =14        
	xlMajorGridlines              =15        
	xlMinorGridlines              =16        
	xlAxisTitle                   =17        
	xlUpBars                      =18        
	xlPlotArea                    =19        
	xlDownBars                    =20        
	xlAxis                        =21        
	xlSeriesLines                 =22        
	xlFloor                       =23        
	xlLegend                      =24        
	xlHiLoLines                   =25        
	xlDropLines                   =26        
	xlRadarAxisLabels             =27        
	xlNothing                     =28        
	xlLeaderLines                 =29        
	xlDisplayUnitLabel            =30        
	xlPivotChartFieldButton       =31        
	xlPivotChartDropZone          =32        
	xlPivotChartExpandEntireFieldButton=33        
	xlPivotChartCollapseEntireFieldButton=34        

class XlChartOrientation(Enum):
	xlDownward                    =-4170     
	xlHorizontal                  =-4128     
	xlUpward                      =-4171     
	xlVertical                    =-4166     

class XlChartPictureType(Enum):
	xlStackScale                  =3         
	xlStack                       =2         
	xlStretch                     =1         

class XlChartSplitType(Enum):
	xlSplitByPosition             =1         
	xlSplitByPercentValue         =3         
	xlSplitByCustomSplit          =4         
	xlSplitByValue                =2         

class XlChartType(Enum):
	xlColumnClustered             =51        
	xlColumnStacked               =52        
	xlColumnStacked100            =53        
	xl3DColumnClustered           =54        
	xl3DColumnStacked             =55        
	xl3DColumnStacked100          =56        
	xlBarClustered                =57        
	xlBarStacked                  =58        
	xlBarStacked100               =59        
	xl3DBarClustered              =60        
	xl3DBarStacked                =61        
	xl3DBarStacked100             =62        
	xlLineStacked                 =63        
	xlLineStacked100              =64        
	xlLineMarkers                 =65        
	xlLineMarkersStacked          =66        
	xlLineMarkersStacked100       =67        
	xlPieOfPie                    =68        
	xlPieExploded                 =69        
	xl3DPieExploded               =70        
	xlBarOfPie                    =71        
	xlXYScatterSmooth             =72        
	xlXYScatterSmoothNoMarkers    =73        
	xlXYScatterLines              =74        
	xlXYScatterLinesNoMarkers     =75        
	xlAreaStacked                 =76        
	xlAreaStacked100              =77        
	xl3DAreaStacked               =78        
	xl3DAreaStacked100            =79        
	xlDoughnutExploded            =80        
	xlRadarMarkers                =81        
	xlRadarFilled                 =82        
	xlSurface                     =83        
	xlSurfaceWireframe            =84        
	xlSurfaceTopView              =85        
	xlSurfaceTopViewWireframe     =86        
	xlBubble                      =15        
	xlBubble3DEffect              =87        
	xlStockHLC                    =88        
	xlStockOHLC                   =89        
	xlStockVHLC                   =90        
	xlStockVOHLC                  =91        
	xlCylinderColClustered        =92        
	xlCylinderColStacked          =93        
	xlCylinderColStacked100       =94        
	xlCylinderBarClustered        =95        
	xlCylinderBarStacked          =96        
	xlCylinderBarStacked100       =97        
	xlCylinderCol                 =98        
	xlConeColClustered            =99        
	xlConeColStacked              =100       
	xlConeColStacked100           =101       
	xlConeBarClustered            =102       
	xlConeBarStacked              =103       
	xlConeBarStacked100           =104       
	xlConeCol                     =105       
	xlPyramidColClustered         =106       
	xlPyramidColStacked           =107       
	xlPyramidColStacked100        =108       
	xlPyramidBarClustered         =109       
	xlPyramidBarStacked           =110       
	xlPyramidBarStacked100        =111       
	xlPyramidCol                  =112       
	xl3DColumn                    =-4100     
	xlLine                        =4         
	xl3DLine                      =-4101     
	xl3DPie                       =-4102     
	xlPie                         =5         
	xlXYScatter                   =-4169     
	xl3DArea                      =-4098     
	xlArea                        =1         
	xlDoughnut                    =-4120     
	xlRadar                       =-4151     
	xlCombo                       =-4152     
	xlComboColumnClusteredLine    =113       
	xlComboColumnClusteredLineSecondaryAxis=114       
	xlComboAreaStackedColumnClustered=115       
	xlOtherCombinations           =116       
	xlSuggestedChart              =-2        
	xlTreemap                     =117       
	xlHistogram                   =118       
	xlWaterfall                   =119       
	xlSunburst                    =120       
	xlBoxwhisker                  =121       
	xlPareto                      =122       
	xlFunnel                      =123       
	xlColumnClusteredEx           =124       
	xlColumnStackedEx             =125       
	xlColumnStacked100Ex          =126       
	xlLineEx                      =127       
	xlLineStackedEx               =128       
	xlLineStacked100Ex            =129       
	xlPieEx                       =130       
	xlDoughnutEx                  =131       
	xlBarClusteredEx              =132       
	xlBarStackedEx                =133       
	xlBarStacked100Ex             =134       
	xlAreaEx                      =135       
	xlAreaStackedEx               =136       
	xlAreaStacked100Ex            =137       
	xlXYScatterEx                 =138       
	xlBubbleEx                    =139       
	xlRegionMap                   =140       

class XlColorIndex(Enum):
	xlColorIndexAutomatic         =-4105     
	xlColorIndexNone              =-4142     

class XlConstants(Enum):
	xlAutomatic                   =-4105     
	xlCombination                 =-4111     
	xlCustom                      =-4114     
	xlBar                         =2         
	xlColumn                      =3         
	xl3DBar                       =-4099     
	xl3DSurface                   =-4103     
	xlDefaultAutoFormat           =-1        
	xlNone                        =-4142     
	xlAbove                       =0         
	xlBelow                       =1         
	xlBoth                        =1         
	xlBottom                      =-4107     
	xlCenter                      =-4108     
	xlChecker                     =9         
	xlCircle                      =8         
	xlCorner                      =2         
	xlCrissCross                  =16        
	xlCross                       =4         
	xlDiamond                     =2         
	xlDistributed                 =-4117     
	xlFill                        =5         
	xlFixedValue                  =1         
	xlGeneral                     =1         
	xlGray16                      =17        
	xlGray25                      =-4124     
	xlGray50                      =-4125     
	xlGray75                      =-4126     
	xlGray8                       =18        
	xlGrid                        =15        
	xlHigh                        =-4127     
	xlInside                      =2         
	xlJustify                     =-4130     
	xlLeft                        =-4131     
	xlLightDown                   =13        
	xlLightHorizontal             =11        
	xlLightUp                     =14        
	xlLightVertical               =12        
	xlLow                         =-4134     
	xlMaximum                     =2         
	xlMinimum                     =4         
	xlMinusValues                 =3         
	xlNextToAxis                  =4         
	xlOpaque                      =3         
	xlOutside                     =3         
	xlPercent                     =2         
	xlPlus                        =9         
	xlPlusValues                  =2         
	xlRight                       =-4152     
	xlScale                       =3         
	xlSemiGray75                  =10        
	xlShowLabel                   =4         
	xlShowLabelAndPercent         =5         
	xlShowPercent                 =3         
	xlShowValue                   =2         
	xlSingle                      =2         
	xlSolid                       =1         
	xlSquare                      =1         
	xlStar                        =5         
	xlStError                     =4         
	xlTop                         =-4160     
	xlTransparent                 =2         
	xlTriangle                    =3         

class XlDataLabelPosition(Enum):
	xlLabelPositionCenter         =-4108     
	xlLabelPositionAbove          =0         
	xlLabelPositionBelow          =1         
	xlLabelPositionLeft           =-4131     
	xlLabelPositionRight          =-4152     
	xlLabelPositionOutsideEnd     =2         
	xlLabelPositionInsideEnd      =3         
	xlLabelPositionInsideBase     =4         
	xlLabelPositionBestFit        =5         
	xlLabelPositionMixed          =6         
	xlLabelPositionCustom         =7         

class XlDataLabelsType(Enum):
	xlDataLabelsShowNone          =-4142     
	xlDataLabelsShowValue         =2         
	xlDataLabelsShowPercent       =3         
	xlDataLabelsShowLabel         =4         
	xlDataLabelsShowLabelAndPercent=5         
	xlDataLabelsShowBubbleSizes   =6         

class XlDisplayBlanksAs(Enum):
	xlInterpolated                =3         
	xlNotPlotted                  =1         
	xlZero                        =2         

class XlDisplayUnit(Enum):
	xlHundreds                    =-2        
	xlThousands                   =-3        
	xlTenThousands                =-4        
	xlHundredThousands            =-5        
	xlMillions                    =-6        
	xlTenMillions                 =-7        
	xlHundredMillions             =-8        
	xlThousandMillions            =-9        
	xlMillionMillions             =-10       
	xlDisplayUnitCustom           =-4114     
	xlDisplayUnitNone             =-4142     

class XlEndStyleCap(Enum):
	xlCap                         =1         
	xlNoCap                       =2         

class XlErrorBarDirection(Enum):
	xlChartX                      =-4168     
	xlChartY                      =1         

class XlErrorBarInclude(Enum):
	xlErrorBarIncludeBoth         =1         
	xlErrorBarIncludeMinusValues  =3         
	xlErrorBarIncludeNone         =-4142     
	xlErrorBarIncludePlusValues   =2         

class XlErrorBarType(Enum):
	xlErrorBarTypeCustom          =-4114     
	xlErrorBarTypeFixedValue      =1         
	xlErrorBarTypePercent         =2         
	xlErrorBarTypeStDev           =-4155     
	xlErrorBarTypeStError         =4         

class XlGeoMappingLevel(Enum):
	xlGeoMappingLevelAutomatic    =0         
	xlGeoMappingLevelDataOnly     =1         
	xlGeoMappingLevelPostalCode   =2         
	xlGeoMappingLevelCounty       =3         
	xlGeoMappingLevelState        =4         
	xlGeoMappingLevelCountryRegion=5         
	xlGeoMappingLevelCountryRegionList=6         
	xlGeoMappingLevelWorld        =7         

class XlGeoProjectionType(Enum):
	xlGeoProjectionTypeAutomatic  =0         
	xlGeoProjectionTypeMercator   =1         
	xlGeoProjectionTypeMiller     =2         
	xlGeoProjectionTypeAlbers     =3         
	xlGeoProjectionTypeRobinson   =4         

class XlGradientStopPositionType(Enum):
	xlGradientStopPositionTypeExtremeValue=0         
	xlGradientStopPositionTypeNumber=1         
	xlGradientStopPositionTypePercent=2         

class XlHAlign(Enum):
	xlHAlignCenter                =-4108     
	xlHAlignCenterAcrossSelection =7         
	xlHAlignDistributed           =-4117     
	xlHAlignFill                  =5         
	xlHAlignGeneral               =1         
	xlHAlignJustify               =-4130     
	xlHAlignLeft                  =-4131     
	xlHAlignRight                 =-4152     

class XlLegendPosition(Enum):
	xlLegendPositionBottom        =-4107     
	xlLegendPositionCorner        =2         
	xlLegendPositionLeft          =-4131     
	xlLegendPositionRight         =-4152     
	xlLegendPositionTop           =-4160     
	xlLegendPositionCustom        =-4161     

class XlMarkerStyle(Enum):
	xlMarkerStyleAutomatic        =-4105     
	xlMarkerStyleCircle           =8         
	xlMarkerStyleDash             =-4115     
	xlMarkerStyleDiamond          =2         
	xlMarkerStyleDot              =-4118     
	xlMarkerStyleNone             =-4142     
	xlMarkerStylePicture          =-4147     
	xlMarkerStylePlus             =9         
	xlMarkerStyleSquare           =1         
	xlMarkerStyleStar             =5         
	xlMarkerStyleTriangle         =3         
	xlMarkerStyleX                =-4168     

class XlParentDataLabelOptions(Enum):
	xlParentDataLabelOptionsNone  =0         
	xlParentDataLabelOptionsBanner=1         
	xlParentDataLabelOptionsOverlapping=2         

class XlPieSliceIndex(Enum):
	xlOuterCounterClockwisePoint  =1         
	xlOuterCenterPoint            =2         
	xlOuterClockwisePoint         =3         
	xlMidClockwiseRadiusPoint     =4         
	xlCenterPoint                 =5         
	xlMidCounterClockwiseRadiusPoint=6         
	xlInnerClockwisePoint         =7         
	xlInnerCenterPoint            =8         
	xlInnerCounterClockwisePoint  =9         

class XlPieSliceLocation(Enum):
	xlHorizontalCoordinate        =1         
	xlVerticalCoordinate          =2         

class XlPivotFieldOrientation(Enum):
	xlColumnField                 =2         
	xlDataField                   =4         
	xlHidden                      =0         
	xlPageField                   =3         
	xlRowField                    =1         

class XlReadingOrder(Enum):
	xlContext                     =-5002     
	xlLTR                         =-5003     
	xlRTL                         =-5004     

class XlRegionLabelOptions(Enum):
	xlRegionLabelOptionsNone      =0         
	xlRegionLabelOptionsBestFitOnly=1         
	xlRegionLabelOptionsShowAll   =2         

class XlRowCol(Enum):
	xlColumns                     =2         
	xlRows                        =1         

class XlScaleType(Enum):
	xlScaleLinear                 =-4132     
	xlScaleLogarithmic            =-4133     

class XlSeriesColorGradientStyle(Enum):
	xlSeriesColorGradientStyleSequential=0         
	xlSeriesColorGradientStyleDiverging=1         

class XlSeriesNameLevel(Enum):
	xlSeriesNameLevelNone         =-3        
	xlSeriesNameLevelCustom       =-2        
	xlSeriesNameLevelAll          =-1        

class XlSizeRepresents(Enum):
	xlSizeIsWidth                 =2         
	xlSizeIsArea                  =1         

class XlTickLabelOrientation(Enum):
	xlTickLabelOrientationAutomatic=-4105     
	xlTickLabelOrientationDownward=-4170     
	xlTickLabelOrientationHorizontal=-4128     
	xlTickLabelOrientationUpward  =-4171     
	xlTickLabelOrientationVertical=-4166     

class XlTickLabelPosition(Enum):
	xlTickLabelPositionHigh       =-4127     
	xlTickLabelPositionLow        =-4134     
	xlTickLabelPositionNextToAxis =4         
	xlTickLabelPositionNone       =-4142     

class XlTickMark(Enum):
	xlTickMarkCross               =4         
	xlTickMarkInside              =2         
	xlTickMarkNone                =-4142     
	xlTickMarkOutside             =3         

class XlTimeUnit(Enum):
	xlDays                        =0         
	xlMonths                      =1         
	xlYears                       =2         

class XlTrendlineType(Enum):
	xlExponential                 =5         
	xlLinear                      =-4132     
	xlLogarithmic                 =-4133     
	xlMovingAvg                   =6         
	xlPolynomial                  =3         
	xlPower                       =4         

class XlUnderlineStyle(Enum):
	xlUnderlineStyleDouble        =-4119     
	xlUnderlineStyleDoubleAccounting=5         
	xlUnderlineStyleNone          =-4142     
	xlUnderlineStyleSingle        =2         
	xlUnderlineStyleSingleAccounting=4         

class XlVAlign(Enum):
	xlVAlignBottom                =-4107     
	xlVAlignCenter                =-4108     
	xlVAlignDistributed           =-4117     
	xlVAlignJustify               =-4130     
	xlVAlignTop                   =-4160     

class XlValueSortOrder(Enum):
	xlValueNone                   =0         
	xlValueAscending              =1         
	xlValueDescending             =2         


class Adjustments(typing.Protocol):

	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> float:
		...
	# The method SetItem is actually a property, but must be used as a method to correctly pass the arguments
	def SetItem(self, Index: int=defaultNamedNotOptArg, arg1: float=defaultUnnamedArg) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> float:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class AnswerWizard(typing.Protocol):

	def ClearFileList(self) -> None:
		...
	def ResetFileList(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
		# Method 'Files' returns object of type 'AnswerWizardFiles'
	Files: AnswerWizardFiles
	Parent: typing.Any
	def __iter__(self):
		...

class AnswerWizardFiles(typing.Protocol):

	def Add(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	def Delete(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Assistant(typing.Protocol):

	def ActivateWizard(self, WizardID: int=defaultNamedNotOptArg, act: MsoWizardActType=defaultNamedNotOptArg, Animation: typing.Any=defaultNamedOptArg) -> None:
		...
	def DoAlert(self, bstrAlertTitle: str=defaultNamedNotOptArg, bstrAlertText: str=defaultNamedNotOptArg, alb: MsoAlertButtonType=defaultNamedNotOptArg, alc: MsoAlertIconType=defaultNamedNotOptArg
			, ald: MsoAlertDefaultType=defaultNamedNotOptArg, alq: MsoAlertCancelType=defaultNamedNotOptArg, varfSysAlert: bool=defaultNamedNotOptArg) -> int:
		...
	def EndWizard(self, WizardID: int=defaultNamedNotOptArg, varfSuccess: bool=defaultNamedNotOptArg, Animation: typing.Any=defaultNamedOptArg) -> None:
		...
	def Help(self) -> None:
		...
	def Move(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> None:
		...
	def ResetTips(self) -> None:
		...
	def StartWizard(self, On: bool=defaultNamedNotOptArg, Callback: str=defaultNamedNotOptArg, PrivateX: int=defaultNamedNotOptArg, Animation: typing.Any=defaultNamedOptArg
			, CustomTeaser: typing.Any=defaultNamedOptArg, Top: typing.Any=defaultNamedOptArg, Left: typing.Any=defaultNamedOptArg, Bottom: typing.Any=defaultNamedOptArg, Right: typing.Any=defaultNamedOptArg) -> int:
		...
		# Method 'Animation' returns enumeration of type 'MsoAnimationType'
	Animation: MsoAnimationType
	Application: typing.Any
	AssistWithAlerts: typing.Any
	AssistWithHelp: typing.Any
	AssistWithWizards: typing.Any
		# Method 'BalloonError' returns enumeration of type 'MsoBalloonErrorType'
	BalloonError: MsoBalloonErrorType
	Creator: typing.Any
	FeatureTips: typing.Any
	FileName: typing.Any
	GuessHelp: typing.Any
	HighPriorityTips: typing.Any
	Item: typing.Any
	KeyboardShortcutTips: typing.Any
	Left: typing.Any
	MouseTips: typing.Any
	MoveWhenInTheWay: typing.Any
	Name: typing.Any
		# Method 'NewBalloon' returns object of type 'Balloon'
	NewBalloon: Balloon
	On: typing.Any
	Parent: typing.Any
	Reduced: typing.Any
	SearchWhenProgramming: typing.Any
	Sounds: typing.Any
	TipOfDay: typing.Any
	Top: typing.Any
	Visible: typing.Any
	# Default property for this class is 'Item'
	def __call__(self):
		...
	def __iter__(self):
		...

class Axes(typing.Protocol):

	# Result is of type IMsoAxis
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Type: XlAxisType=defaultNamedNotOptArg, AxisGroup: XlAxisGroup=1) -> IMsoAxis:
		...
	# Result is of type IMsoAxis
	# The method _Default is actually a property, but must be used as a method to correctly pass the arguments
	def _Default(self, Type: XlAxisType=defaultNamedNotOptArg, AxisGroup: XlAxisGroup=1) -> IMsoAxis:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	_NewEnum: typing.Any
	# Default method for this class is '_Default'
	def __call__(self, Type: XlAxisType=defaultNamedNotOptArg, AxisGroup: XlAxisGroup=1) -> IMsoAxis:
		...
	def __iter__(self):
		...
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Balloon(typing.Protocol):

	def Close(self) -> None:
		...
	def SetAvoidRectangle(self, Left: int=defaultNamedNotOptArg, Top: int=defaultNamedNotOptArg, Right: int=defaultNamedNotOptArg, Bottom: int=defaultNamedNotOptArg) -> None:
		...
	# Result is of type MsoBalloonButtonType
	def Show(self) -> MsoBalloonButtonType:
		...
		# Method 'Animation' returns enumeration of type 'MsoAnimationType'
	Animation: MsoAnimationType
	Application: typing.Any
		# Method 'BalloonType' returns enumeration of type 'MsoBalloonType'
	BalloonType: MsoBalloonType
		# Method 'Button' returns enumeration of type 'MsoButtonSetType'
	Button: MsoButtonSetType
	Callback: typing.Any
	Checkboxes: typing.Any
	Creator: typing.Any
	Heading: typing.Any
		# Method 'Icon' returns enumeration of type 'MsoIconType'
	Icon: MsoIconType
	Labels: typing.Any
		# Method 'Mode' returns enumeration of type 'MsoModeType'
	Mode: MsoModeType
	Name: typing.Any
	Parent: typing.Any
	Private: typing.Any
	Text: typing.Any
	def __iter__(self):
		...

class BalloonCheckbox(typing.Protocol):

	Application: typing.Any
	Checked: typing.Any
	Creator: typing.Any
	Item: typing.Any
	Name: typing.Any
	Parent: typing.Any
	Text: typing.Any
	# Default property for this class is 'Item'
	def __call__(self):
		...
	def __iter__(self):
		...

class BalloonCheckboxes(typing.Protocol):

	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> Dispatch:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Name: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> Dispatch:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class BalloonLabel(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Item: typing.Any
	Name: typing.Any
	Parent: typing.Any
	Text: typing.Any
	# Default property for this class is 'Item'
	def __call__(self):
		...
	def __iter__(self):
		...

class BalloonLabels(typing.Protocol):

	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> Dispatch:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Name: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> Dispatch:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class BulletFormat2(typing.Protocol):

	def Picture(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Character: typing.Any
	Creator: typing.Any
		# Method 'Font' returns object of type 'Font2'
	Font: Font2
	Number: typing.Any
	Parent: typing.Any
	RelativeSize: typing.Any
	StartValue: typing.Any
		# Method 'Style' returns enumeration of type 'MsoNumberedBulletStyle'
	Style: MsoNumberedBulletStyle
		# Method 'Type' returns enumeration of type 'MsoBulletType'
	Type: MsoBulletType
		# Method 'UseTextColor' returns enumeration of type 'MsoTriState'
	UseTextColor: MsoTriState
		# Method 'UseTextFont' returns enumeration of type 'MsoTriState'
	UseTextFont: MsoTriState
		# Method 'Visible' returns enumeration of type 'MsoTriState'
	Visible: MsoTriState
	def __iter__(self):
		...

class COMAddIn(typing.Protocol):

	Application: typing.Any
	Connect: typing.Any
	Creator: typing.Any
	Description: typing.Any
	Guid: typing.Any
	Object: typing.Any
	Parent: typing.Any
	ProgId: typing.Any
	# Default property for this class is 'Description'
	def __call__(self):
		...
	def __iter__(self):
		...

class COMAddIns(typing.Protocol):

	# Result is of type COMAddIn
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> COMAddIn:
		...
	def SetAppModal(self, varfModal: bool=defaultNamedNotOptArg) -> None:
		...
	def Update(self) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> COMAddIn:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class CalloutFormat(typing.Protocol):

	def AutomaticLength(self) -> None:
		...
	def CustomDrop(self, Drop: float=defaultNamedNotOptArg) -> None:
		...
	def CustomLength(self, Length: float=defaultNamedNotOptArg) -> None:
		...
	def PresetDrop(self, DropType: MsoCalloutDropType=defaultNamedNotOptArg) -> None:
		...
		# Method 'Accent' returns enumeration of type 'MsoTriState'
	Accent: MsoTriState
		# Method 'Angle' returns enumeration of type 'MsoCalloutAngleType'
	Angle: MsoCalloutAngleType
	Application: typing.Any
		# Method 'AutoAttach' returns enumeration of type 'MsoTriState'
	AutoAttach: MsoTriState
		# Method 'AutoLength' returns enumeration of type 'MsoTriState'
	AutoLength: MsoTriState
		# Method 'Border' returns enumeration of type 'MsoTriState'
	Border: MsoTriState
	Creator: typing.Any
	Drop: typing.Any
		# Method 'DropType' returns enumeration of type 'MsoCalloutDropType'
	DropType: MsoCalloutDropType
	Gap: typing.Any
	Length: typing.Any
	Parent: typing.Any
		# Method 'Type' returns enumeration of type 'MsoCalloutType'
	Type: MsoCalloutType
	def __iter__(self):
		...

class CanvasShapes(typing.Protocol):

	# Result is of type Shape
	def AddCallout(self, Type: MsoCalloutType=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddConnector(self, Type: MsoConnectorType=defaultNamedNotOptArg, BeginX: float=defaultNamedNotOptArg, BeginY: float=defaultNamedNotOptArg, EndX: float=defaultNamedNotOptArg
			, EndY: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddCurve(self, SafeArrayOfPoints: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddLabel(self, Orientation: MsoTextOrientation=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddLine(self, BeginX: float=defaultNamedNotOptArg, BeginY: float=defaultNamedNotOptArg, EndX: float=defaultNamedNotOptArg, EndY: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddPicture(self, FileName: str=defaultNamedNotOptArg, LinkToFile: MsoTriState=defaultNamedNotOptArg, SaveWithDocument: MsoTriState=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg
			, Top: float=defaultNamedNotOptArg, Width: float=-1.0, Height: float=-1.0) -> Shape:
		...
	# Result is of type Shape
	def AddPolyline(self, SafeArrayOfPoints: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddShape(self, Type: MsoAutoShapeType=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddTextEffect(self, PresetTextEffect: MsoPresetTextEffect=defaultNamedNotOptArg, Text: str=defaultNamedNotOptArg, FontName: str=defaultNamedNotOptArg, FontSize: float=defaultNamedNotOptArg
			, FontBold: MsoTriState=defaultNamedNotOptArg, FontItalic: MsoTriState=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddTextbox(self, Orientation: MsoTextOrientation=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type FreeformBuilder
	def BuildFreeform(self, EditingType: MsoEditingType=defaultNamedNotOptArg, X1: float=defaultNamedNotOptArg, Y1: float=defaultNamedNotOptArg) -> FreeformBuilder:
		...
	# Result is of type Shape
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type ShapeRange
	def Range(self, Index: typing.Any=defaultNamedNotOptArg) -> ShapeRange:
		...
	def SelectAll(self) -> None:
		...
	Application: typing.Any
		# Method 'Background' returns object of type 'Shape'
	Background: Shape
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class CategoryCollection(typing.Protocol):

	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ChartColorFormat(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	RGB: typing.Any
	SchemeColor: typing.Any
	Type: typing.Any
	_Default: typing.Any
	# Default property for this class is '_Default'
	def __call__(self):
		...
	def __iter__(self):
		...

class ChartFillFormat(typing.Protocol):

	def OneColorGradient(self, Style: int=defaultNamedNotOptArg, Variant: int=defaultNamedNotOptArg, Degree: float=defaultNamedNotOptArg) -> None:
		...
	def Patterned(self, Pattern: int=defaultNamedNotOptArg) -> None:
		...
	def PresetGradient(self, Style: int=defaultNamedNotOptArg, Variant: int=defaultNamedNotOptArg, PresetGradientType: int=defaultNamedNotOptArg) -> None:
		...
	def PresetTextured(self, PresetTexture: int=defaultNamedNotOptArg) -> None:
		...
	def Solid(self) -> None:
		...
	def TwoColorGradient(self, Style: int=defaultNamedNotOptArg, Variant: int=defaultNamedNotOptArg) -> None:
		...
	def UserPicture(self, PictureFile: typing.Any=defaultNamedNotOptArg, PictureFormat: typing.Any=defaultNamedNotOptArg, PictureStackUnit: typing.Any=defaultNamedNotOptArg, PicturePlacement: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def UserTextured(self, TextureFile: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'BackColor' returns object of type 'ChartColorFormat'
	BackColor: ChartColorFormat
	Creator: typing.Any
		# Method 'ForeColor' returns object of type 'ChartColorFormat'
	ForeColor: ChartColorFormat
	GradientColorType: typing.Any
	GradientDegree: typing.Any
	GradientStyle: typing.Any
	GradientVariant: typing.Any
	Parent: typing.Any
	Pattern: typing.Any
	PresetGradientType: typing.Any
	PresetTexture: typing.Any
	TextureName: typing.Any
	TextureType: typing.Any
	Type: typing.Any
	Visible: typing.Any
	def __iter__(self):
		...

class ChartFont(typing.Protocol):

	Application: typing.Any
	Background: typing.Any
	Bold: typing.Any
	Color: typing.Any
	ColorIndex: typing.Any
	Creator: typing.Any
	FontStyle: typing.Any
	Italic: typing.Any
	Name: typing.Any
	OutlineFont: typing.Any
	Parent: typing.Any
	Shadow: typing.Any
	Size: typing.Any
	StrikeThrough: typing.Any
	Subscript: typing.Any
	Superscript: typing.Any
	Underline: typing.Any
	def __iter__(self):
		...

class ChartGroups(typing.Protocol):

	# Result is of type IMsoChartGroup
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> IMsoChartGroup:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ChartPoint(typing.Protocol):

	Application: typing.Any
	ApplyPictToEnd: typing.Any
	ApplyPictToFront: typing.Any
	ApplyPictToSides: typing.Any
	Border: typing.Any
	Creator: typing.Any
	DataLabel: typing.Any
	Explosion: typing.Any
	Fill: typing.Any
	Format: typing.Any
	Has3DEffect: typing.Any
	HasDataLabel: typing.Any
	Height: typing.Any
	Interior: typing.Any
	InvertIfNegative: typing.Any
	IsTotal: typing.Any
	Left: typing.Any
	MarkerBackgroundColor: typing.Any
	MarkerBackgroundColorIndex: typing.Any
	MarkerForegroundColor: typing.Any
	MarkerForegroundColorIndex: typing.Any
	MarkerSize: typing.Any
	MarkerStyle: typing.Any
	Name: typing.Any
	Parent: typing.Any
	PictureType: typing.Any
	PictureUnit: typing.Any
	PictureUnit2: typing.Any
	SecondaryPlot: typing.Any
	Shadow: typing.Any
	Top: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class ColorFormat(typing.Protocol):

	Application: typing.Any
	Brightness: typing.Any
	Creator: typing.Any
		# Method 'ObjectThemeColor' returns enumeration of type 'MsoThemeColorIndex'
	ObjectThemeColor: MsoThemeColorIndex
	Parent: typing.Any
	RGB: typing.Any
	SchemeColor: typing.Any
	TintAndShade: typing.Any
		# Method 'Type' returns enumeration of type 'MsoColorType'
	Type: MsoColorType
	# Default property for this class is 'RGB'
	def __call__(self):
		...
	def __iter__(self):
		...

class CommandBar(typing.Protocol):

	def Delete(self) -> None:
		...
	# Result is of type CommandBarControl
	def FindControl(self, Type: typing.Any=defaultNamedOptArg, Id: typing.Any=defaultNamedOptArg, Tag: typing.Any=defaultNamedOptArg, Visible: typing.Any=defaultNamedOptArg
			, Recursive: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	# The method GetaccDefaultAction is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccDescription is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDescription(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelp is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelp(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelpTopic is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelpTopic(self, pszHelpFile: str=pythoncom.Missing, varChild: typing.Any=defaultNamedOptArg) -> int:
		...
	# The method GetaccKeyboardShortcut is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccKeyboardShortcut(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccName(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccRole is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccRole(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccState is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccState(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccValue(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	def Reset(self) -> None:
		...
	# The method SetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccName(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccValue(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	def ShowPopup(self, x: typing.Any=defaultNamedOptArg, y: typing.Any=defaultNamedOptArg) -> None:
		...
	# The method accChild is actually a property, but must be used as a method to correctly pass the arguments
	def accChild(self, varChild: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	def accDoDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accHitTest(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> typing.Any:
		...
	def accLocation(self, pxLeft: int=pythoncom.Missing, pyTop: int=pythoncom.Missing, pcxWidth: int=pythoncom.Missing, pcyHeight: int=pythoncom.Missing
			, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accNavigate(self, navDir: int=defaultNamedNotOptArg, varStart: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def accSelect(self, flagsSelect: int=defaultNamedNotOptArg, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	AdaptiveMenu: typing.Any
	Application: typing.Any
	BuiltIn: typing.Any
	Context: typing.Any
		# Method 'Controls' returns object of type 'CommandBarControls'
	Controls: CommandBarControls
	Creator: typing.Any
	Enabled: typing.Any
	Height: typing.Any
	Id: typing.Any
	Index: typing.Any
	InstanceId: typing.Any
	InstanceIdPtr: typing.Any
	Left: typing.Any
	Name: typing.Any
	NameLocal: typing.Any
	Parent: typing.Any
		# Method 'Position' returns enumeration of type 'MsoBarPosition'
	Position: MsoBarPosition
		# Method 'Protection' returns enumeration of type 'MsoBarProtection'
	Protection: MsoBarProtection
	RowIndex: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'MsoBarType'
	Type: MsoBarType
	Visible: typing.Any
	Width: typing.Any
	accChildCount: typing.Any
	accDefaultAction: typing.Any
	accDescription: typing.Any
	accFocus: typing.Any
	accHelp: typing.Any
	accHelpTopic: typing.Any
	accKeyboardShortcut: typing.Any
	accName: typing.Any
	accParent: typing.Any
	accRole: typing.Any
	accSelection: typing.Any
	accState: typing.Any
	accValue: typing.Any
	def __iter__(self):
		...

class CommandBarControl(typing.Protocol):

	# Result is of type CommandBarControl
	def Copy(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def Delete(self, Temporary: typing.Any=defaultNamedOptArg) -> None:
		...
	def Execute(self) -> None:
		...
	# The method GetaccDefaultAction is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccDescription is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDescription(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelp is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelp(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelpTopic is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelpTopic(self, pszHelpFile: str=pythoncom.Missing, varChild: typing.Any=defaultNamedOptArg) -> int:
		...
	# The method GetaccKeyboardShortcut is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccKeyboardShortcut(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccName(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccRole is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccRole(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccState is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccState(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccValue(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# Result is of type CommandBarControl
	def Move(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def Reserved1(self) -> None:
		...
	def Reserved2(self) -> None:
		...
	def Reserved3(self) -> None:
		...
	def Reserved4(self) -> None:
		...
	def Reserved5(self) -> None:
		...
	def Reserved6(self) -> None:
		...
	def Reserved7(self) -> None:
		...
	def Reset(self) -> None:
		...
	def SetFocus(self) -> None:
		...
	# The method SetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccName(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccValue(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method accChild is actually a property, but must be used as a method to correctly pass the arguments
	def accChild(self, varChild: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	def accDoDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accHitTest(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> typing.Any:
		...
	def accLocation(self, pxLeft: int=pythoncom.Missing, pyTop: int=pythoncom.Missing, pcxWidth: int=pythoncom.Missing, pcyHeight: int=pythoncom.Missing
			, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accNavigate(self, navDir: int=defaultNamedNotOptArg, varStart: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def accSelect(self, flagsSelect: int=defaultNamedNotOptArg, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
	BeginGroup: typing.Any
	BuiltIn: typing.Any
	Caption: typing.Any
	Control: typing.Any
	Creator: typing.Any
	DescriptionText: typing.Any
	Enabled: typing.Any
	Height: typing.Any
	HelpContextId: typing.Any
	HelpFile: typing.Any
	Id: typing.Any
	Index: typing.Any
	InstanceId: typing.Any
	IsPriorityDropped: typing.Any
	Left: typing.Any
		# Method 'OLEUsage' returns enumeration of type 'MsoControlOLEUsage'
	OLEUsage: MsoControlOLEUsage
	OnAction: typing.Any
	Parameter: typing.Any
		# Method 'Parent' returns object of type 'CommandBar'
	Parent: CommandBar
	Priority: typing.Any
	Tag: typing.Any
	TooltipText: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'MsoControlType'
	Type: MsoControlType
	Visible: typing.Any
	Width: typing.Any
	accChildCount: typing.Any
	accDefaultAction: typing.Any
	accDescription: typing.Any
	accFocus: typing.Any
	accHelp: typing.Any
	accHelpTopic: typing.Any
	accKeyboardShortcut: typing.Any
	accName: typing.Any
	accParent: typing.Any
	accRole: typing.Any
	accSelection: typing.Any
	accState: typing.Any
	accValue: typing.Any
	def __iter__(self):
		...

class CommandBarControls(typing.Protocol):

	# Result is of type CommandBarControl
	def Add(self, Type: typing.Any=defaultNamedOptArg, Id: typing.Any=defaultNamedOptArg, Parameter: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg
			, Temporary: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	# Result is of type CommandBarControl
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> CommandBarControl:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
		# Method 'Parent' returns object of type 'CommandBar'
	Parent: CommandBar
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> CommandBarControl:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class CommandBarPopup(typing.Protocol):

	# Result is of type CommandBarControl
	def Copy(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def Delete(self, Temporary: typing.Any=defaultNamedOptArg) -> None:
		...
	def Execute(self) -> None:
		...
	# The method GetaccDefaultAction is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccDescription is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDescription(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelp is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelp(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelpTopic is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelpTopic(self, pszHelpFile: str=pythoncom.Missing, varChild: typing.Any=defaultNamedOptArg) -> int:
		...
	# The method GetaccKeyboardShortcut is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccKeyboardShortcut(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccName(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccRole is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccRole(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccState is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccState(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccValue(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# Result is of type CommandBarControl
	def Move(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def Reserved1(self) -> None:
		...
	def Reserved2(self) -> None:
		...
	def Reserved3(self) -> None:
		...
	def Reserved4(self) -> None:
		...
	def Reserved5(self) -> None:
		...
	def Reserved6(self) -> None:
		...
	def Reserved7(self) -> None:
		...
	def Reset(self) -> None:
		...
	def SetFocus(self) -> None:
		...
	# The method SetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccName(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccValue(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method accChild is actually a property, but must be used as a method to correctly pass the arguments
	def accChild(self, varChild: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	def accDoDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accHitTest(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> typing.Any:
		...
	def accLocation(self, pxLeft: int=pythoncom.Missing, pyTop: int=pythoncom.Missing, pcxWidth: int=pythoncom.Missing, pcyHeight: int=pythoncom.Missing
			, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accNavigate(self, navDir: int=defaultNamedNotOptArg, varStart: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def accSelect(self, flagsSelect: int=defaultNamedNotOptArg, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
	BeginGroup: typing.Any
	BuiltIn: typing.Any
	Caption: typing.Any
		# Method 'CommandBar' returns object of type 'CommandBar'
	CommandBar: CommandBar
	Control: typing.Any
		# Method 'Controls' returns object of type 'CommandBarControls'
	Controls: CommandBarControls
	Creator: typing.Any
	DescriptionText: typing.Any
	Enabled: typing.Any
	Height: typing.Any
	HelpContextId: typing.Any
	HelpFile: typing.Any
	Id: typing.Any
	Index: typing.Any
	InstanceId: typing.Any
	InstanceIdPtr: typing.Any
	IsPriorityDropped: typing.Any
	Left: typing.Any
		# Method 'OLEMenuGroup' returns enumeration of type 'MsoOLEMenuGroup'
	OLEMenuGroup: MsoOLEMenuGroup
		# Method 'OLEUsage' returns enumeration of type 'MsoControlOLEUsage'
	OLEUsage: MsoControlOLEUsage
	OnAction: typing.Any
	Parameter: typing.Any
		# Method 'Parent' returns object of type 'CommandBar'
	Parent: CommandBar
	Priority: typing.Any
	Tag: typing.Any
	TooltipText: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'MsoControlType'
	Type: MsoControlType
	Visible: typing.Any
	Width: typing.Any
	accChildCount: typing.Any
	accDefaultAction: typing.Any
	accDescription: typing.Any
	accFocus: typing.Any
	accHelp: typing.Any
	accHelpTopic: typing.Any
	accKeyboardShortcut: typing.Any
	accName: typing.Any
	accParent: typing.Any
	accRole: typing.Any
	accSelection: typing.Any
	accState: typing.Any
	accValue: typing.Any
	def __iter__(self):
		...

class ConnectorFormat(typing.Protocol):

	def BeginConnect(self, ConnectedShape: Shape=defaultNamedNotOptArg, ConnectionSite: int=defaultNamedNotOptArg) -> None:
		...
	def BeginDisconnect(self) -> None:
		...
	def EndConnect(self, ConnectedShape: Shape=defaultNamedNotOptArg, ConnectionSite: int=defaultNamedNotOptArg) -> None:
		...
	def EndDisconnect(self) -> None:
		...
	Application: typing.Any
		# Method 'BeginConnected' returns enumeration of type 'MsoTriState'
	BeginConnected: MsoTriState
		# Method 'BeginConnectedShape' returns object of type 'Shape'
	BeginConnectedShape: Shape
	BeginConnectionSite: typing.Any
	Creator: typing.Any
		# Method 'EndConnected' returns enumeration of type 'MsoTriState'
	EndConnected: MsoTriState
		# Method 'EndConnectedShape' returns object of type 'Shape'
	EndConnectedShape: Shape
	EndConnectionSite: typing.Any
	Parent: typing.Any
		# Method 'Type' returns enumeration of type 'MsoConnectorType'
	Type: MsoConnectorType
	def __iter__(self):
		...

class ContactCard(typing.Protocol):

	def Close(self) -> None:
		...
	def Show(self, CardStyle: MsoContactCardStyle=defaultNamedNotOptArg, RectangleLeft: int=defaultNamedNotOptArg, RectangleRight: int=defaultNamedNotOptArg, RectangleTop: int=defaultNamedNotOptArg
			, RectangleBottom: int=defaultNamedNotOptArg, HorizontalPosition: int=defaultNamedNotOptArg, ShowWithDelay: bool=False) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	def __iter__(self):
		...

class Crop(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	PictureHeight: typing.Any
	PictureOffsetX: typing.Any
	PictureOffsetY: typing.Any
	PictureWidth: typing.Any
	ShapeHeight: typing.Any
	ShapeLeft: typing.Any
	ShapeTop: typing.Any
	ShapeWidth: typing.Any
	# Default property for this class is 'PictureOffsetX'
	def __call__(self):
		...
	def __iter__(self):
		...

class CustomTaskPaneEvents(typing.Protocol):

	def DockPositionStateChange(self, CustomTaskPaneInst: _CustomTaskPane=defaultNamedNotOptArg) -> None:
		...
	def VisibleStateChange(self, CustomTaskPaneInst: _CustomTaskPane=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class CustomXMLNode(typing.Protocol):

	def AppendChildNode(self, Name: str='', NamespaceURI: str='', NodeType: MsoCustomXMLNodeType=1, NodeValue: str='') -> None:
		...
	def AppendChildSubtree(self, XML: str=defaultNamedNotOptArg) -> None:
		...
	def Delete(self) -> None:
		...
	def HasChildNodes(self) -> bool:
		...
	def InsertNodeBefore(self, Name: str='', NamespaceURI: str='', NodeType: MsoCustomXMLNodeType=1, NodeValue: str=''
			, NextSibling: CustomXMLNode=0) -> None:
		...
	def InsertSubtreeBefore(self, XML: str=defaultNamedNotOptArg, NextSibling: CustomXMLNode=0) -> None:
		...
	def RemoveChild(self, Child: CustomXMLNode=defaultNamedNotOptArg) -> None:
		...
	def ReplaceChildNode(self, OldNode: CustomXMLNode=defaultNamedNotOptArg, Name: str='', NamespaceURI: str='', NodeType: MsoCustomXMLNodeType=1
			, NodeValue: str='') -> None:
		...
	def ReplaceChildSubtree(self, XML: str=defaultNamedNotOptArg, OldNode: CustomXMLNode=defaultNamedNotOptArg) -> None:
		...
	# Result is of type CustomXMLNodes
	def SelectNodes(self, XPath: str=defaultNamedNotOptArg) -> CustomXMLNodes:
		...
	# Result is of type CustomXMLNode
	def SelectSingleNode(self, XPath: str=defaultNamedNotOptArg) -> CustomXMLNode:
		...
	Application: typing.Any
		# Method 'Attributes' returns object of type 'CustomXMLNodes'
	Attributes: CustomXMLNodes
	BaseName: typing.Any
		# Method 'ChildNodes' returns object of type 'CustomXMLNodes'
	ChildNodes: CustomXMLNodes
	Creator: typing.Any
		# Method 'FirstChild' returns object of type 'CustomXMLNode'
	FirstChild: CustomXMLNode
		# Method 'LastChild' returns object of type 'CustomXMLNode'
	LastChild: CustomXMLNode
	NamespaceURI: typing.Any
		# Method 'NextSibling' returns object of type 'CustomXMLNode'
	NextSibling: CustomXMLNode
		# Method 'NodeType' returns enumeration of type 'MsoCustomXMLNodeType'
	NodeType: MsoCustomXMLNodeType
	NodeValue: typing.Any
	OwnerDocument: typing.Any
		# Method 'OwnerPart' returns object of type 'CustomXMLPart'
	OwnerPart: CustomXMLPart
	Parent: typing.Any
		# Method 'ParentNode' returns object of type 'CustomXMLNode'
	ParentNode: CustomXMLNode
		# Method 'PreviousSibling' returns object of type 'CustomXMLNode'
	PreviousSibling: CustomXMLNode
	Text: typing.Any
	XML: typing.Any
	XPath: typing.Any
	def __iter__(self):
		...

class CustomXMLNodes(typing.Protocol):

	# Result is of type CustomXMLNode
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> CustomXMLNode:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> CustomXMLNode:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class CustomXMLPrefixMapping(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	NamespaceURI: typing.Any
	Parent: typing.Any
	Prefix: typing.Any
	def __iter__(self):
		...

class CustomXMLPrefixMappings(typing.Protocol):

	def AddNamespace(self, Prefix: str=defaultNamedNotOptArg, NamespaceURI: str=defaultNamedNotOptArg) -> None:
		...
	# Result is of type CustomXMLPrefixMapping
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> CustomXMLPrefixMapping:
		...
	def LookupNamespace(self, Prefix: str=defaultNamedNotOptArg) -> str:
		...
	def LookupPrefix(self, NamespaceURI: str=defaultNamedNotOptArg) -> str:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> CustomXMLPrefixMapping:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class CustomXMLSchema(typing.Protocol):

	def Delete(self) -> None:
		...
	def Reload(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	Location: typing.Any
	NamespaceURI: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class CustomXMLValidationError(typing.Protocol):

	def Delete(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	ErrorCode: typing.Any
	Name: typing.Any
		# Method 'Node' returns object of type 'CustomXMLNode'
	Node: CustomXMLNode
	Parent: typing.Any
	Text: typing.Any
		# Method 'Type' returns enumeration of type 'MsoCustomXMLValidationErrorType'
	Type: MsoCustomXMLValidationErrorType
	def __iter__(self):
		...

class CustomXMLValidationErrors(typing.Protocol):

	def Add(self, Node: CustomXMLNode=defaultNamedNotOptArg, ErrorName: str=defaultNamedNotOptArg, ErrorText: str='', ClearedOnUpdate: bool=True) -> None:
		...
	# Result is of type CustomXMLValidationError
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> CustomXMLValidationError:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> CustomXMLValidationError:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class DataPrivacyOptions(typing.Protocol):

	def GetPrivacyOptionSetting(self, PrivacyOption: MsoPrivacyOptionSetting=defaultNamedNotOptArg) -> bool:
		...
	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
		# Method 'SendTelemetryOption' returns enumeration of type 'MsoTelemetryConsentLevel'
	SendTelemetryOption: MsoTelemetryConsentLevel
	def __iter__(self):
		...

class DiagramNode(typing.Protocol):

	# Result is of type DiagramNode
	def AddNode(self, Pos: MsoRelativeNodePosition=2, NodeType: MsoDiagramNodeType=1) -> DiagramNode:
		...
	# Result is of type DiagramNode
	def CloneNode(self, CopyChildren: bool=defaultNamedNotOptArg, TargetNode: DiagramNode=defaultNamedNotOptArg, Pos: MsoRelativeNodePosition=2) -> DiagramNode:
		...
	def Delete(self) -> None:
		...
	def MoveNode(self, TargetNode: DiagramNode=defaultNamedNotOptArg, Pos: MsoRelativeNodePosition=defaultNamedNotOptArg) -> None:
		...
	# Result is of type DiagramNode
	def NextNode(self) -> DiagramNode:
		...
	# Result is of type DiagramNode
	def PrevNode(self) -> DiagramNode:
		...
	def ReplaceNode(self, TargetNode: DiagramNode=defaultNamedNotOptArg) -> None:
		...
	def SwapNode(self, TargetNode: DiagramNode=defaultNamedNotOptArg, SwapChildren: bool=True) -> None:
		...
	def TransferChildren(self, ReceivingNode: DiagramNode=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'Children' returns object of type 'DiagramNodeChildren'
	Children: DiagramNodeChildren
	Creator: typing.Any
		# Method 'Diagram' returns object of type 'IMsoDiagram'
	Diagram: IMsoDiagram
		# Method 'Layout' returns enumeration of type 'MsoOrgChartLayoutType'
	Layout: MsoOrgChartLayoutType
	Parent: typing.Any
		# Method 'Root' returns object of type 'DiagramNode'
	Root: DiagramNode
		# Method 'Shape' returns object of type 'Shape'
	Shape: Shape
		# Method 'TextShape' returns object of type 'Shape'
	TextShape: Shape
	def __iter__(self):
		...

class DiagramNodeChildren(typing.Protocol):

	# Result is of type DiagramNode
	def AddNode(self, Index: typing.Any=-1, NodeType: MsoDiagramNodeType=1) -> DiagramNode:
		...
	# Result is of type DiagramNode
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> DiagramNode:
		...
	def SelectAll(self) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
		# Method 'FirstChild' returns object of type 'DiagramNode'
	FirstChild: DiagramNode
		# Method 'LastChild' returns object of type 'DiagramNode'
	LastChild: DiagramNode
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> DiagramNode:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class DiagramNodes(typing.Protocol):

	# Result is of type DiagramNode
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> DiagramNode:
		...
	def SelectAll(self) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> DiagramNode:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class DocumentInspector(typing.Protocol):

	def Fix(self, Status: MsoDocInspectorStatus=pythoncom.Missing, Results: str=pythoncom.Missing) -> None:
		...
	def Inspect(self, Status: MsoDocInspectorStatus=pythoncom.Missing, Results: str=pythoncom.Missing) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	Description: typing.Any
	Name: typing.Any
	Parent: typing.Any
	# Default property for this class is 'Name'
	def __call__(self):
		...
	def __iter__(self):
		...

class DocumentInspectors(typing.Protocol):

	# Result is of type DocumentInspector
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> DocumentInspector:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> DocumentInspector:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class DocumentLibraryVersion(typing.Protocol):

	def Delete(self) -> None:
		...
	def Open(self) -> Dispatch:
		...
	def Restore(self) -> Dispatch:
		...
	Application: typing.Any
	Comments: typing.Any
	Creator: typing.Any
	Index: typing.Any
	Modified: typing.Any
	ModifiedBy: typing.Any
	Parent: typing.Any
	# Default property for this class is 'Modified'
	def __call__(self):
		...
	def __iter__(self):
		...

class DocumentLibraryVersions(typing.Protocol):

	# Result is of type DocumentLibraryVersion
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, lIndex: int=defaultNamedNotOptArg) -> DocumentLibraryVersion:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	IsVersioningEnabled: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, lIndex: int=defaultNamedNotOptArg) -> DocumentLibraryVersion:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class DocumentProperties(typing.Protocol):

	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class DocumentProperty(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	LinkSource: typing.Any
	LinkToContent: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class EffectParameter(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Name: typing.Any
	Value: typing.Any
	# Default property for this class is 'Name'
	def __call__(self):
		...
	def __iter__(self):
		...

class EffectParameters(typing.Protocol):

	# Result is of type EffectParameter
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> EffectParameter:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> EffectParameter:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class EncryptionProvider(typing.Protocol):

	def Authenticate(self, ParentWindow: typing.Any=defaultNamedNotOptArg, EncryptionData: typing.Any=defaultNamedNotOptArg, PermissionsMask: int=pythoncom.Missing) -> int:
		...
	def CloneSession(self, SessionHandle: int=defaultNamedNotOptArg) -> int:
		...
	def DecryptStream(self, SessionHandle: int=defaultNamedNotOptArg, StreamName: str=defaultNamedNotOptArg, EncryptedStream: typing.Any=defaultNamedNotOptArg, UnencryptedStream: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def EncryptStream(self, SessionHandle: int=defaultNamedNotOptArg, StreamName: str=defaultNamedNotOptArg, UnencryptedStream: typing.Any=defaultNamedNotOptArg, EncryptedStream: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def EndSession(self, SessionHandle: int=defaultNamedNotOptArg) -> None:
		...
	def GetProviderDetail(self, encprovdet: EncryptionProviderDetail=defaultNamedNotOptArg) -> typing.Any:
		...
	def NewSession(self, ParentWindow: typing.Any=defaultNamedNotOptArg) -> int:
		...
	def Save(self, SessionHandle: int=defaultNamedNotOptArg, EncryptionData: typing.Any=defaultNamedNotOptArg) -> int:
		...
	def ShowSettings(self, SessionHandle: int=defaultNamedNotOptArg, ParentWindow: typing.Any=defaultNamedNotOptArg, ReadOnly: bool=defaultNamedNotOptArg, Remove: bool=pythoncom.Missing) -> None:
		...
	def __iter__(self):
		...

class FileDialog(typing.Protocol):

	def Execute(self) -> None:
		...
	def Show(self) -> int:
		...
	AllowMultiSelect: typing.Any
	Application: typing.Any
	ButtonName: typing.Any
	Creator: typing.Any
		# Method 'DialogType' returns enumeration of type 'MsoFileDialogType'
	DialogType: MsoFileDialogType
	FilterIndex: typing.Any
		# Method 'Filters' returns object of type 'FileDialogFilters'
	Filters: FileDialogFilters
	InitialFileName: typing.Any
		# Method 'InitialView' returns enumeration of type 'MsoFileDialogView'
	InitialView: MsoFileDialogView
	Item: typing.Any
	Parent: typing.Any
		# Method 'SelectedItems' returns object of type 'FileDialogSelectedItems'
	SelectedItems: FileDialogSelectedItems
	Title: typing.Any
	# Default property for this class is 'Item'
	def __call__(self):
		...
	def __iter__(self):
		...

class FileDialogFilter(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Description: typing.Any
	Extensions: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class FileDialogFilters(typing.Protocol):

	# Result is of type FileDialogFilter
	def Add(self, Description: str=defaultNamedNotOptArg, Extensions: str=defaultNamedNotOptArg, Position: typing.Any=defaultNamedOptArg) -> FileDialogFilter:
		...
	def Clear(self) -> None:
		...
	def Delete(self, filter: typing.Any=defaultNamedOptArg) -> None:
		...
	# Result is of type FileDialogFilter
	def Item(self, Index: int=defaultNamedNotOptArg) -> FileDialogFilter:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> FileDialogFilter:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class FileDialogSelectedItems(typing.Protocol):

	def Item(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class FileSearch(typing.Protocol):

	def Execute(self, SortBy: MsoSortBy=1, SortOrder: MsoSortOrder=1, AlwaysAccurate: bool=True) -> int:
		...
	def NewSearch(self) -> None:
		...
	def RefreshScopes(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	FileName: typing.Any
		# Method 'FileType' returns enumeration of type 'MsoFileType'
	FileType: MsoFileType
		# Method 'FileTypes' returns object of type 'FileTypes'
	FileTypes: FileTypes
		# Method 'FoundFiles' returns object of type 'FoundFiles'
	FoundFiles: FoundFiles
		# Method 'LastModified' returns enumeration of type 'MsoLastModified'
	LastModified: MsoLastModified
	LookIn: typing.Any
	MatchAllWordForms: typing.Any
	MatchTextExactly: typing.Any
		# Method 'PropertyTests' returns object of type 'PropertyTests'
	PropertyTests: PropertyTests
		# Method 'SearchFolders' returns object of type 'SearchFolders'
	SearchFolders: SearchFolders
		# Method 'SearchScopes' returns object of type 'SearchScopes'
	SearchScopes: SearchScopes
	SearchSubFolders: typing.Any
	TextOrProperty: typing.Any
	def __iter__(self):
		...

class FileTypes(typing.Protocol):

	def Add(self, FileType: MsoFileType=defaultNamedNotOptArg) -> None:
		...
	# Result is of type MsoFileType
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> MsoFileType:
		...
	def Remove(self, Index: int=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> MsoFileType:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class FillFormat(typing.Protocol):

	def Background(self) -> None:
		...
	def OneColorGradient(self, Style: MsoGradientStyle=defaultNamedNotOptArg, Variant: int=defaultNamedNotOptArg, Degree: float=defaultNamedNotOptArg) -> None:
		...
	def Patterned(self, Pattern: MsoPatternType=defaultNamedNotOptArg) -> None:
		...
	def PresetGradient(self, Style: MsoGradientStyle=defaultNamedNotOptArg, Variant: int=defaultNamedNotOptArg, PresetGradientType: MsoPresetGradientType=defaultNamedNotOptArg) -> None:
		...
	def PresetTextured(self, PresetTexture: MsoPresetTexture=defaultNamedNotOptArg) -> None:
		...
	def Solid(self) -> None:
		...
	def TwoColorGradient(self, Style: MsoGradientStyle=defaultNamedNotOptArg, Variant: int=defaultNamedNotOptArg) -> None:
		...
	def UserPicture(self, PictureFile: str=defaultNamedNotOptArg) -> None:
		...
	def UserTextured(self, TextureFile: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'BackColor' returns object of type 'ColorFormat'
	BackColor: ColorFormat
	Creator: typing.Any
		# Method 'ForeColor' returns object of type 'ColorFormat'
	ForeColor: ColorFormat
	GradientAngle: typing.Any
		# Method 'GradientColorType' returns enumeration of type 'MsoGradientColorType'
	GradientColorType: MsoGradientColorType
	GradientDegree: typing.Any
		# Method 'GradientStops' returns object of type 'GradientStops'
	GradientStops: GradientStops
		# Method 'GradientStyle' returns enumeration of type 'MsoGradientStyle'
	GradientStyle: MsoGradientStyle
	GradientVariant: typing.Any
	Parent: typing.Any
		# Method 'Pattern' returns enumeration of type 'MsoPatternType'
	Pattern: MsoPatternType
		# Method 'PictureEffects' returns object of type 'PictureEffects'
	PictureEffects: PictureEffects
		# Method 'PresetGradientType' returns enumeration of type 'MsoPresetGradientType'
	PresetGradientType: MsoPresetGradientType
		# Method 'PresetTexture' returns enumeration of type 'MsoPresetTexture'
	PresetTexture: MsoPresetTexture
		# Method 'RotateWithObject' returns enumeration of type 'MsoTriState'
	RotateWithObject: MsoTriState
		# Method 'TextureAlignment' returns enumeration of type 'MsoTextureAlignment'
	TextureAlignment: MsoTextureAlignment
	TextureHorizontalScale: typing.Any
	TextureName: typing.Any
	TextureOffsetX: typing.Any
	TextureOffsetY: typing.Any
		# Method 'TextureTile' returns enumeration of type 'MsoTriState'
	TextureTile: MsoTriState
		# Method 'TextureType' returns enumeration of type 'MsoTextureType'
	TextureType: MsoTextureType
	TextureVerticalScale: typing.Any
	Transparency: typing.Any
		# Method 'Type' returns enumeration of type 'MsoFillType'
	Type: MsoFillType
		# Method 'Visible' returns enumeration of type 'MsoTriState'
	Visible: MsoTriState
	def __iter__(self):
		...

class Font2(typing.Protocol):

		# Method 'Allcaps' returns enumeration of type 'MsoTriState'
	Allcaps: MsoTriState
	Application: typing.Any
		# Method 'AutorotateNumbers' returns enumeration of type 'MsoTriState'
	AutorotateNumbers: MsoTriState
	BaselineOffset: typing.Any
		# Method 'Bold' returns enumeration of type 'MsoTriState'
	Bold: MsoTriState
		# Method 'Caps' returns enumeration of type 'MsoTextCaps'
	Caps: MsoTextCaps
	Creator: typing.Any
		# Method 'DoubleStrikeThrough' returns enumeration of type 'MsoTriState'
	DoubleStrikeThrough: MsoTriState
		# Method 'Embeddable' returns enumeration of type 'MsoTriState'
	Embeddable: MsoTriState
		# Method 'Embedded' returns enumeration of type 'MsoTriState'
	Embedded: MsoTriState
		# Method 'Equalize' returns enumeration of type 'MsoTriState'
	Equalize: MsoTriState
		# Method 'Fill' returns object of type 'FillFormat'
	Fill: FillFormat
		# Method 'Glow' returns object of type 'GlowFormat'
	Glow: GlowFormat
		# Method 'Highlight' returns object of type 'ColorFormat'
	Highlight: ColorFormat
		# Method 'Italic' returns enumeration of type 'MsoTriState'
	Italic: MsoTriState
	Kerning: typing.Any
		# Method 'Line' returns object of type 'LineFormat'
	Line: LineFormat
	Name: typing.Any
	NameAscii: typing.Any
	NameComplexScript: typing.Any
	NameFarEast: typing.Any
	NameOther: typing.Any
	Parent: typing.Any
		# Method 'Reflection' returns object of type 'ReflectionFormat'
	Reflection: ReflectionFormat
		# Method 'Shadow' returns object of type 'ShadowFormat'
	Shadow: ShadowFormat
	Size: typing.Any
		# Method 'Smallcaps' returns enumeration of type 'MsoTriState'
	Smallcaps: MsoTriState
		# Method 'SoftEdgeFormat' returns enumeration of type 'MsoSoftEdgeType'
	SoftEdgeFormat: MsoSoftEdgeType
	Spacing: typing.Any
		# Method 'Strike' returns enumeration of type 'MsoTextStrike'
	Strike: MsoTextStrike
		# Method 'StrikeThrough' returns enumeration of type 'MsoTriState'
	StrikeThrough: MsoTriState
		# Method 'Subscript' returns enumeration of type 'MsoTriState'
	Subscript: MsoTriState
		# Method 'Superscript' returns enumeration of type 'MsoTriState'
	Superscript: MsoTriState
		# Method 'UnderlineColor' returns object of type 'ColorFormat'
	UnderlineColor: ColorFormat
		# Method 'UnderlineStyle' returns enumeration of type 'MsoTextUnderlineType'
	UnderlineStyle: MsoTextUnderlineType
		# Method 'WordArtformat' returns enumeration of type 'MsoPresetTextEffect'
	WordArtformat: MsoPresetTextEffect
	def __iter__(self):
		...

class FoundFiles(typing.Protocol):

	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class FreeformBuilder(typing.Protocol):

	def AddNodes(self, SegmentType: MsoSegmentType=defaultNamedNotOptArg, EditingType: MsoEditingType=defaultNamedNotOptArg, X1: float=defaultNamedNotOptArg, Y1: float=defaultNamedNotOptArg
			, X2: float=0.0, Y2: float=0.0, X3: float=0.0, Y3: float=0.0) -> None:
		...
	# Result is of type Shape
	def ConvertToShape(self) -> Shape:
		...
	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class FullSeriesCollection(typing.Protocol):

	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class GlowFormat(typing.Protocol):

	Application: typing.Any
		# Method 'Color' returns object of type 'ColorFormat'
	Color: ColorFormat
	Creator: typing.Any
	Radius: typing.Any
	Transparency: typing.Any
	def __iter__(self):
		...

class GradientStop(typing.Protocol):

	Application: typing.Any
		# Method 'Color' returns object of type 'ColorFormat'
	Color: ColorFormat
	Creator: typing.Any
	Position: typing.Any
	Transparency: typing.Any
	def __iter__(self):
		...

class GradientStops(typing.Protocol):

	def Delete(self, Index: int=-1) -> None:
		...
	def Insert(self, RGB: int=defaultNamedNotOptArg, Position: float=defaultNamedNotOptArg, Transparency: float=0.0, Index: int=-1) -> None:
		...
	def Insert2(self, RGB: int=defaultNamedNotOptArg, Position: float=defaultNamedNotOptArg, Transparency: float=0.0, Index: int=-1
			, Brightness: float=0.0) -> None:
		...
	# Result is of type GradientStop
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> GradientStop:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> GradientStop:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class GridLines(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class GroupShapes(typing.Protocol):

	# Result is of type Shape
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type ShapeRange
	def Range(self, Index: typing.Any=defaultNamedNotOptArg) -> ShapeRange:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class HTMLProject(typing.Protocol):

	def Open(self, OpenKind: MsoHTMLProjectOpen=0) -> None:
		...
	def RefreshDocument(self, Refresh: bool=True) -> None:
		...
	def RefreshProject(self, Refresh: bool=True) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
		# Method 'HTMLProjectItems' returns object of type 'HTMLProjectItems'
	HTMLProjectItems: HTMLProjectItems
	Parent: typing.Any
		# Method 'State' returns enumeration of type 'MsoHTMLProjectState'
	State: MsoHTMLProjectState
	# Default property for this class is 'State'
	def __call__(self):
		...
	def __iter__(self):
		...

class HTMLProjectItem(typing.Protocol):

	def LoadFromFile(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	def Open(self, OpenKind: MsoHTMLProjectOpen=0) -> None:
		...
	def SaveCopyAs(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	IsOpen: typing.Any
	Name: typing.Any
	Parent: typing.Any
	Text: typing.Any
	# Default property for this class is 'Name'
	def __call__(self):
		...
	def __iter__(self):
		...

class HTMLProjectItems(typing.Protocol):

	# Result is of type HTMLProjectItem
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> HTMLProjectItem:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> HTMLProjectItem:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class IAccessible(typing.Protocol):

	# The method GetaccDefaultAction is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccDescription is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDescription(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelp is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelp(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelpTopic is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelpTopic(self, pszHelpFile: str=pythoncom.Missing, varChild: typing.Any=defaultNamedOptArg) -> int:
		...
	# The method GetaccKeyboardShortcut is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccKeyboardShortcut(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccName(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccRole is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccRole(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccState is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccState(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccValue(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method SetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccName(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccValue(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method accChild is actually a property, but must be used as a method to correctly pass the arguments
	def accChild(self, varChild: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	def accDoDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accHitTest(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> typing.Any:
		...
	def accLocation(self, pxLeft: int=pythoncom.Missing, pyTop: int=pythoncom.Missing, pcxWidth: int=pythoncom.Missing, pcyHeight: int=pythoncom.Missing
			, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accNavigate(self, navDir: int=defaultNamedNotOptArg, varStart: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def accSelect(self, flagsSelect: int=defaultNamedNotOptArg, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	accChildCount: typing.Any
	accDefaultAction: typing.Any
	accDescription: typing.Any
	accFocus: typing.Any
	accHelp: typing.Any
	accHelpTopic: typing.Any
	accKeyboardShortcut: typing.Any
	accName: typing.Any
	accParent: typing.Any
	accRole: typing.Any
	accSelection: typing.Any
	accState: typing.Any
	accValue: typing.Any
	def __iter__(self):
		...

class IAssistance(typing.Protocol):

	def ClearDefaultContext(self, HelpId: str=defaultNamedNotOptArg) -> None:
		'ClearDefaultContext Method'
		...
	def SearchHelp(self, Query: str=defaultNamedNotOptArg, Scope: str='') -> None:
		'SearchHelp Method'
		...
	def SetDefaultContext(self, HelpId: str=defaultNamedNotOptArg) -> None:
		'SetDefaultContext Method'
		...
	def ShowHelp(self, HelpId: str='', Scope: str='') -> None:
		'ShowHelp Method'
		...
	def __iter__(self):
		...

class IBlogExtensibility(typing.Protocol):

	def BlogProviderProperties(self, BlogProvider: str=pythoncom.Missing, FriendlyName: str=pythoncom.Missing, CategorySupport: MsoBlogCategorySupport=pythoncom.Missing, Padding: bool=pythoncom.Missing) -> None:
		...
	def GetCategories(self, Account: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, Document: Dispatch=defaultNamedNotOptArg, Categories: typing.List[str]=pythoncom.Missing) -> None:
		...
	def GetRecentPosts(self, Account: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, Document: Dispatch=defaultNamedNotOptArg, PostTitles: typing.List[str]=pythoncom.Missing
			, PostDates: typing.List[str]=pythoncom.Missing, PostIDs: typing.List[str]=pythoncom.Missing) -> None:
		...
	def GetUserBlogs(self, Account: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, Document: Dispatch=defaultNamedNotOptArg, BlogNames: typing.List[str]=pythoncom.Missing
			, BlogIDs: typing.List[str]=pythoncom.Missing, BlogURLs: typing.List[str]=pythoncom.Missing) -> None:
		...
	def Open(self, Account: str=defaultNamedNotOptArg, PostID: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, xHTML: str=pythoncom.Missing
			, Title: str=pythoncom.Missing, DatePosted: str=pythoncom.Missing, Categories: typing.List[str]=pythoncom.Missing) -> None:
		...
	def PublishPost(self, Account: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, Document: Dispatch=defaultNamedNotOptArg, xHTML: str=defaultNamedNotOptArg
			, Title: str=defaultNamedNotOptArg, DateTime: str=defaultNamedNotOptArg, Categories: typing.List[str]=defaultNamedNotOptArg, Draft: bool=defaultNamedNotOptArg, PostID: str=pythoncom.Missing
			, PublishMessage: str=pythoncom.Missing) -> None:
		...
	def RepublishPost(self, Account: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, Document: Dispatch=defaultNamedNotOptArg, PostID: str=defaultNamedNotOptArg
			, xHTML: str=defaultNamedNotOptArg, Title: str=defaultNamedNotOptArg, DateTime: str=defaultNamedNotOptArg, Categories: typing.List[str]=defaultNamedNotOptArg, Draft: bool=defaultNamedNotOptArg
			, PublishMessage: str=pythoncom.Missing) -> None:
		...
	def SetupBlogAccount(self, Account: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, Document: Dispatch=defaultNamedNotOptArg, NewAccount: bool=defaultNamedNotOptArg
			, ShowPictureUI: bool=pythoncom.Missing) -> None:
		...
	def __iter__(self):
		...

class IBlogPictureExtensibility(typing.Protocol):

	def BlogPictureProviderProperties(self, BlogPictureProvider: str=pythoncom.Missing, FriendlyName: str=pythoncom.Missing) -> None:
		...
	def CreatePictureAccount(self, Account: str=defaultNamedNotOptArg, BlogProvider: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, Document: Dispatch=defaultNamedNotOptArg) -> None:
		...
	def PublishPicture(self, Account: str=defaultNamedNotOptArg, ParentWindow: int=defaultNamedNotOptArg, Document: Dispatch=defaultNamedNotOptArg, Image: typing.Any=defaultNamedNotOptArg
			, PictureURI: str=pythoncom.Missing, ImageType: int=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class ICTPFactory(typing.Protocol):

	# Result is of type _CustomTaskPane
	def CreateCTP(self, CTPAxID: str=defaultNamedNotOptArg, CTPTitle: str=defaultNamedNotOptArg, CTPParentWindow: typing.Any=defaultNamedOptArg) -> _CustomTaskPane:
		...
	def __iter__(self):
		...

class ICommandBarButtonEvents(typing.Protocol):

	def Click(self, Ctrl: CommandBarButton=defaultNamedNotOptArg, CancelDefault: bool=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class ICommandBarComboBoxEvents(typing.Protocol):

	def Change(self, Ctrl: CommandBarComboBox=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class ICommandBarsEvents(typing.Protocol):

	def OnUpdate(self) -> None:
		...
	def __iter__(self):
		...

class ICustomTaskPaneConsumer(typing.Protocol):

	def CTPFactoryAvailable(self, CTPFactoryInst: ICTPFactory=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class ICustomXMLPartEvents(typing.Protocol):

	def NodeAfterDelete(self, OldNode: CustomXMLNode=defaultNamedNotOptArg, OldParentNode: CustomXMLNode=defaultNamedNotOptArg, OldNextSibling: CustomXMLNode=defaultNamedNotOptArg, InUndoRedo: bool=defaultNamedNotOptArg) -> None:
		...
	def NodeAfterInsert(self, NewNode: CustomXMLNode=defaultNamedNotOptArg, InUndoRedo: bool=defaultNamedNotOptArg) -> None:
		...
	def NodeAfterReplace(self, OldNode: CustomXMLNode=defaultNamedNotOptArg, NewNode: CustomXMLNode=defaultNamedNotOptArg, InUndoRedo: bool=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class ICustomXMLPartsEvents(typing.Protocol):

	def PartAfterAdd(self, NewPart: CustomXMLPart=defaultNamedNotOptArg) -> None:
		...
	def PartAfterLoad(self, Part: CustomXMLPart=defaultNamedNotOptArg) -> None:
		...
	def PartBeforeDelete(self, OldPart: CustomXMLPart=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class IFind(typing.Protocol):

	def Delete(self, bstrQueryName: str=defaultNamedNotOptArg) -> None:
		...
	def Execute(self) -> None:
		...
	def Load(self, bstrQueryName: str=defaultNamedNotOptArg) -> None:
		...
	def Save(self, bstrQueryName: str=defaultNamedNotOptArg) -> None:
		...
	def Show(self) -> int:
		...
	Author: typing.Any
	DateCreatedFrom: typing.Any
	DateCreatedTo: typing.Any
	DateSavedFrom: typing.Any
	DateSavedTo: typing.Any
	FileType: typing.Any
	Keywords: typing.Any
		# Method 'ListBy' returns enumeration of type 'MsoFileFindListBy'
	ListBy: MsoFileFindListBy
	MatchCase: typing.Any
	Name: typing.Any
		# Method 'Options' returns enumeration of type 'MsoFileFindOptions'
	Options: MsoFileFindOptions
	PatternMatch: typing.Any
		# Method 'Results' returns object of type 'IFoundFiles'
	Results: IFoundFiles
	SavedBy: typing.Any
	SearchPath: typing.Any
	SelectedFile: typing.Any
		# Method 'SortBy' returns enumeration of type 'MsoFileFindSortBy'
	SortBy: MsoFileFindSortBy
	SubDir: typing.Any
	Subject: typing.Any
	Text: typing.Any
	Title: typing.Any
		# Method 'View' returns enumeration of type 'MsoFileFindView'
	View: MsoFileFindView
	# Default property for this class is 'SearchPath'
	def __call__(self):
		...
	def __iter__(self):
		...

class IFoundFiles(typing.Protocol):

	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	Count: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ILicAgent(typing.Protocol):
	'ILicAgent Interface'

	def AsyncProcessCCRenewalLicenseRequest(self) -> None:
		'method AsyncProcessCCRenewalLicenseRequest'
		...
	def AsyncProcessCCRenewalPriceRequest(self) -> None:
		'method AsyncProcessCCRenewalPriceRequest'
		...
	def AsyncProcessDroppedLicenseRequest(self) -> None:
		'method AsyncProcessDroppedLicenseRequest'
		...
	def AsyncProcessHandshakeRequest(self, bReviseCustInfo: int=defaultNamedNotOptArg) -> None:
		'method AsyncProcessHandshakeRequest'
		...
	def AsyncProcessNewLicenseRequest(self) -> None:
		'method AsyncProcessNewLicenseRequest'
		...
	def AsyncProcessReissueLicenseRequest(self) -> None:
		'method AsyncProcessReissueLicenseRequest'
		...
	def AsyncProcessRetailRenewalLicenseRequest(self) -> None:
		'method AsyncProcessRetailRenewalLicenseRequest'
		...
	def AsyncProcessReviseCustInfoRequest(self) -> None:
		'method AsyncProcessReviseCustInfoRequest'
		...
	def CancelAsyncProcessRequest(self, bIsLicenseRequest: int=defaultNamedNotOptArg) -> None:
		'method CancelAsyncProcessRequest'
		...
	def CheckSystemClock(self) -> int:
		'method CheckSystemClock'
		...
	def DepositConfirmationId(self, bstrVal: str=defaultNamedNotOptArg) -> int:
		'method DepositConfirmationId'
		...
	def DisplaySSLCert(self) -> int:
		'method DisplaySSLCert'
		...
	def GenerateInstallationId(self) -> str:
		'method GenerateInstallationId'
		...
	def GetAddress1(self) -> str:
		'method GetAddress1'
		...
	def GetAddress2(self) -> str:
		'method GetAddress2'
		...
	def GetAsyncProcessReturnCode(self) -> int:
		'method GetAsyncProcessReturnCode'
		...
	def GetBackendErrorMsg(self) -> str:
		'method GetBackendErrorMsg'
		...
	def GetBillingAddress1(self) -> str:
		'method GetBillingAddress1'
		...
	def GetBillingAddress2(self) -> str:
		'method GetBillingAddress2'
		...
	def GetBillingCity(self) -> str:
		'method GetBillingCity'
		...
	def GetBillingCountryCode(self) -> str:
		'method GetBillingCountryCode'
		...
	def GetBillingFirstName(self) -> str:
		'method GetBillingFirstName'
		...
	def GetBillingLastName(self) -> str:
		'method GetBillingLastName'
		...
	def GetBillingPhone(self) -> str:
		'method GetBillingPhone'
		...
	def GetBillingState(self) -> str:
		'method GetBillingState'
		...
	def GetBillingZip(self) -> str:
		'method GetBillingZip'
		...
	def GetCCRenewalExpiryDate(self) -> float:
		'method GetCCRenewalExpiryDate'
		...
	def GetCity(self) -> str:
		'method GetCity'
		...
	def GetCountryCode(self) -> str:
		'method GetCountryCode'
		...
	def GetCountryDesc(self) -> str:
		'method GetCountryDesc'
		...
	def GetCreditCardCode(self, dwIndex: int=defaultNamedNotOptArg) -> str:
		'method GetCreditCardCode'
		...
	def GetCreditCardCount(self) -> int:
		'method GetCreditCardCount'
		...
	def GetCreditCardExpiryMonth(self) -> int:
		'method GetCreditCardExpiryMonth'
		...
	def GetCreditCardExpiryYear(self) -> int:
		'method GetCreditCardExpiryYear'
		...
	def GetCreditCardName(self, dwIndex: int=defaultNamedNotOptArg) -> str:
		'method GetCreditCardName'
		...
	def GetCreditCardNumber(self) -> str:
		'method GetCreditCardNumber'
		...
	def GetCreditCardType(self) -> str:
		'method GetCreditCardType'
		...
	def GetCurrencyDescription(self, dwCurrencyIndex: int=defaultNamedNotOptArg) -> str:
		'method GetCurrencyDescription'
		...
	def GetCurrencyOption(self) -> int:
		'method GetCurrencyOption'
		...
	def GetCurrentExpiryDate(self) -> float:
		'method GetCurrentExpiryDate'
		...
	def GetDisconnectOption(self) -> int:
		'method GetDisconnectOption'
		...
	def GetEmail(self) -> str:
		'method GetEmail'
		...
	def GetEndOfLifeHtmlText(self) -> str:
		'method GetEndOfLifeHtmlText'
		...
	def GetExistingExpiryDate(self) -> float:
		'method GetExistingExpiryDate'
		...
	def GetFirstName(self) -> str:
		'method GetFirstName'
		...
	def GetInvoiceText(self) -> str:
		'method GetInvoiceText'
		...
	def GetIsoLanguage(self) -> int:
		'method GetIsoLanguage'
		...
	def GetLastName(self) -> str:
		'method GetLastName'
		...
	def GetMSOffer(self) -> str:
		'method GetMSOffer'
		...
	def GetMSUpdate(self) -> str:
		'method GetMSUpdate'
		...
	def GetNewExpiryDate(self) -> float:
		'method GetNewExpiryDate'
		...
	def GetOrgName(self) -> str:
		'method GetOrgName'
		...
	def GetOtherOffer(self) -> str:
		'method GetOtherOffer'
		...
	def GetPhone(self) -> str:
		'method GetPhone'
		...
	def GetPriceItemCount(self) -> int:
		'method GetPriceItemCount'
		...
	def GetPriceItemLabel(self, dwIndex: int=defaultNamedNotOptArg) -> str:
		'method GetPriceItemLabel'
		...
	def GetPriceItemValue(self, dwCurrencyIndex: int=defaultNamedNotOptArg, dwIndex: int=defaultNamedNotOptArg) -> str:
		'method GetPriceItemValue'
		...
	def GetState(self) -> str:
		'method GetState'
		...
	def GetVATLabel(self, bstrCountryCode: str=defaultNamedNotOptArg) -> str:
		'method GetVATLabel'
		...
	def GetVATNumber(self) -> str:
		'method GetVATNumber'
		...
	def GetZip(self) -> str:
		'method GetZip'
		...
	def Initialize(self, dwBPC: int=defaultNamedNotOptArg, dwMode: int=defaultNamedNotOptArg, bstrLicSource: str=defaultNamedNotOptArg) -> int:
		'method Initialize'
		...
	def IsCCRenewalCountry(self, bstrCountryCode: str=defaultNamedNotOptArg) -> int:
		'method IsCCRenewalCountry'
		...
	def IsUpgradeAvailable(self) -> int:
		'method IsUpgradeAvailable'
		...
	def SaveBillingInfo(self, bSave: int=defaultNamedNotOptArg) -> int:
		'method SaveBillingInfo'
		...
	def SetAddress1(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetAddress1'
		...
	def SetAddress2(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetAddress2'
		...
	def SetBillingAddress1(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingAddress1'
		...
	def SetBillingAddress2(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingAddress2'
		...
	def SetBillingCity(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingCity'
		...
	def SetBillingCountryCode(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingCountryCode'
		...
	def SetBillingFirstName(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingFirstName'
		...
	def SetBillingLastName(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingLastName'
		...
	def SetBillingPhone(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingPhone'
		...
	def SetBillingState(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingState'
		...
	def SetBillingZip(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetBillingZip'
		...
	def SetCity(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetCity'
		...
	def SetCountryCode(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetCountryCode'
		...
	def SetCountryDesc(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetCountryDesc'
		...
	def SetCreditCardExpiryMonth(self, dwCCMonth: int=defaultNamedNotOptArg) -> None:
		'method SetCreditCardExpiryMonth'
		...
	def SetCreditCardExpiryYear(self, dwCCYear: int=defaultNamedNotOptArg) -> None:
		'method SetCreditCardExpiryYear'
		...
	def SetCreditCardNumber(self, bstrCCNumber: str=defaultNamedNotOptArg) -> None:
		'method SetCreditCardNumber'
		...
	def SetCreditCardType(self, bstrCCCode: str=defaultNamedNotOptArg) -> None:
		'method SetCreditCardType'
		...
	def SetCurrencyOption(self, dwCurrencyOption: int=defaultNamedNotOptArg) -> None:
		'method SetCurrencyOption'
		...
	def SetDisconnectOption(self, bNewVal: int=defaultNamedNotOptArg) -> None:
		'method SetDisconnectOption'
		...
	def SetEmail(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetEmail'
		...
	def SetFirstName(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetFirstName'
		...
	def SetIsoLanguage(self, dwNewVal: int=defaultNamedNotOptArg) -> None:
		'method SetIsoLanguage'
		...
	def SetLastName(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetLastName'
		...
	def SetMSOffer(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetMSOffer'
		...
	def SetMSUpdate(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetMSUpdate'
		...
	def SetOrgName(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetOrgName'
		...
	def SetOtherOffer(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetOtherOffer'
		...
	def SetPhone(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetPhone'
		...
	def SetState(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetState'
		...
	def SetVATNumber(self, bstrVATNumber: str=defaultNamedNotOptArg) -> None:
		'method SetVATNumber'
		...
	def SetZip(self, bstrNewVal: str=defaultNamedNotOptArg) -> None:
		'method SetZip'
		...
	def VerifyCheckDigits(self, bstrCIDIID: str=defaultNamedNotOptArg) -> int:
		'method VerifyCheckDigits'
		...
	def WantUpgrade(self, bWantUpgrade: int=defaultNamedNotOptArg) -> None:
		'method WantUpgrade'
		...
	def __iter__(self):
		...

class ILicValidator(typing.Protocol):

	Products: typing.Any
	Selection: typing.Any
	def __iter__(self):
		...

class ILicWizExternal(typing.Protocol):

	def DepositPidKey(self, bstrKey: str=defaultNamedNotOptArg, fMORW: int=defaultNamedNotOptArg) -> int:
		...
	def DisableVORWReminder(self, BPC: int=defaultNamedNotOptArg) -> None:
		...
	def FormatDate(self, date: float=defaultNamedNotOptArg, pFormat: str='') -> str:
		...
	def GetConnectedState(self) -> int:
		...
	def InternetDisconnect(self) -> None:
		...
	def InvokeDateTimeApplet(self) -> None:
		...
	def MsoAlert(self, bstrText: str=defaultNamedNotOptArg, bstrButtons: str=defaultNamedNotOptArg, bstrIcon: str=defaultNamedNotOptArg) -> int:
		...
	def OpenInDefaultBrowser(self, bstrUrl: str=defaultNamedNotOptArg) -> None:
		...
	def PrintHtmlDocument(self, punkHtmlDoc: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def ResetPID(self) -> None:
		...
	def ResignDpc(self, bstrProductCode: str=defaultNamedNotOptArg) -> None:
		...
	def SaveReceipt(self, bstrReceipt: str=defaultNamedNotOptArg) -> str:
		...
	def SetDialogSize(self, dx: int=defaultNamedNotOptArg, dy: int=defaultNamedNotOptArg) -> None:
		...
	def ShowHelp(self, pvarId: typing.Any=defaultNamedOptArg) -> None:
		...
	def SortSelectOptions(self, pdispSelect: Dispatch=defaultNamedNotOptArg) -> None:
		...
	def Terminate(self) -> None:
		...
	def VerifyClock(self, lMode: int=defaultNamedNotOptArg) -> int:
		...
	def WriteLog(self, bstrMessage: str=defaultNamedNotOptArg) -> None:
		...
	AnimationEnabled: typing.Any
	Context: typing.Any
	CountryInfo: typing.Any
	LicAgent: typing.Any
	OfficeOnTheWebUrl: typing.Any
	Validator: typing.Any
	def __iter__(self):
		...

class IMsoAxis(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	AxisBetweenCategories: typing.Any
		# Method 'AxisGroup' returns enumeration of type 'XlAxisGroup'
	AxisGroup: XlAxisGroup
		# Method 'AxisTitle' returns object of type 'IMsoAxisTitle'
	AxisTitle: IMsoAxisTitle
		# Method 'BaseUnit' returns enumeration of type 'XlTimeUnit'
	BaseUnit: XlTimeUnit
	BaseUnitIsAuto: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	CategoryNames: typing.Any
		# Method 'CategorySortOrder' returns enumeration of type 'XlCategorySortOrder'
	CategorySortOrder: XlCategorySortOrder
		# Method 'CategoryType' returns enumeration of type 'XlCategoryType'
	CategoryType: XlCategoryType
	Creator: typing.Any
		# Method 'Crosses' returns enumeration of type 'XlAxisCrosses'
	Crosses: XlAxisCrosses
	CrossesAt: typing.Any
		# Method 'DisplayUnit' returns enumeration of type 'XlDisplayUnit'
	DisplayUnit: XlDisplayUnit
	DisplayUnitCustom: typing.Any
		# Method 'DisplayUnitLabel' returns object of type 'IMsoDisplayUnitLabel'
	DisplayUnitLabel: IMsoDisplayUnitLabel
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	HasDisplayUnitLabel: typing.Any
	HasMajorGridlines: typing.Any
	HasMinorGridlines: typing.Any
	HasTitle: typing.Any
	Height: typing.Any
	Left: typing.Any
	LogBase: typing.Any
		# Method 'MajorGridlines' returns object of type 'GridLines'
	MajorGridlines: GridLines
		# Method 'MajorTickMark' returns enumeration of type 'XlTickMark'
	MajorTickMark: XlTickMark
	MajorUnit: typing.Any
	MajorUnitIsAuto: typing.Any
		# Method 'MajorUnitScale' returns enumeration of type 'XlTimeUnit'
	MajorUnitScale: XlTimeUnit
	MaximumScale: typing.Any
	MaximumScaleIsAuto: typing.Any
	MinimumScale: typing.Any
	MinimumScaleIsAuto: typing.Any
		# Method 'MinorGridlines' returns object of type 'GridLines'
	MinorGridlines: GridLines
		# Method 'MinorTickMark' returns enumeration of type 'XlTickMark'
	MinorTickMark: XlTickMark
	MinorUnit: typing.Any
	MinorUnitIsAuto: typing.Any
		# Method 'MinorUnitScale' returns enumeration of type 'XlTimeUnit'
	MinorUnitScale: XlTimeUnit
	Name: typing.Any
	Parent: typing.Any
	ReversePlotOrder: typing.Any
		# Method 'ScaleType' returns enumeration of type 'XlScaleType'
	ScaleType: XlScaleType
		# Method 'TickLabelPosition' returns enumeration of type 'XlTickLabelPosition'
	TickLabelPosition: XlTickLabelPosition
	TickLabelSpacing: typing.Any
	TickLabelSpacingIsAuto: typing.Any
		# Method 'TickLabels' returns object of type 'IMsoTickLabels'
	TickLabels: IMsoTickLabels
	TickMarkSpacing: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'XlAxisType'
	Type: XlAxisType
	Width: typing.Any
	def __iter__(self):
		...

class IMsoAxisTitle(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	# Result is of type IMsoCharacters
	# The method GetCharacters is actually a property, but must be used as a method to correctly pass the arguments
	def GetCharacters(self, Start: typing.Any=defaultNamedOptArg, Length: typing.Any=defaultNamedOptArg) -> IMsoCharacters:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	AutoScaleFont: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Caption: typing.Any
		# Method 'Characters' returns object of type 'IMsoCharacters'
	Characters: IMsoCharacters
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Font' returns object of type 'ChartFont'
	Font: ChartFont
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Formula: typing.Any
	FormulaLocal: typing.Any
	FormulaR1C1: typing.Any
	FormulaR1C1Local: typing.Any
	Height: typing.Any
	HorizontalAlignment: typing.Any
	IncludeInLayout: typing.Any
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Left: typing.Any
	Name: typing.Any
	Orientation: typing.Any
	Parent: typing.Any
		# Method 'Position' returns enumeration of type 'XlChartElementPosition'
	Position: XlChartElementPosition
	ReadingOrder: typing.Any
	Shadow: typing.Any
	Text: typing.Any
	Top: typing.Any
	VerticalAlignment: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class IMsoBorder(typing.Protocol):

	Application: typing.Any
	Color: typing.Any
	ColorIndex: typing.Any
	Creator: typing.Any
	LineStyle: typing.Any
	Parent: typing.Any
	Weight: typing.Any
	def __iter__(self):
		...

class IMsoCategory(typing.Protocol):

	IsFiltered: typing.Any
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoCharacters(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	def Insert(self, bstr: str=defaultNamedNotOptArg) -> typing.Any:
		...
	Application: typing.Any
	Caption: typing.Any
	Count: typing.Any
	Creator: typing.Any
		# Method 'Font' returns object of type 'ChartFont'
	Font: ChartFont
	Parent: typing.Any
	PhoneticCharacters: typing.Any
	Text: typing.Any
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class IMsoChart(typing.Protocol):

	def ApplyChartTemplate(self, bstrFileName: str=defaultNamedNotOptArg) -> None:
		...
	def ApplyCustomType(self, ChartType: XlChartType=defaultNamedNotOptArg, TypeName: typing.Any=defaultNamedOptArg) -> None:
		...
	def ApplyDataLabels(self, Type: XlDataLabelsType=2, IMsoLegendKey: typing.Any=defaultNamedOptArg, AutoText: typing.Any=defaultNamedOptArg, HasLeaderLines: typing.Any=defaultNamedOptArg
			, ShowSeriesName: typing.Any=defaultNamedOptArg, ShowCategoryName: typing.Any=defaultNamedOptArg, ShowValue: typing.Any=defaultNamedOptArg, ShowPercentage: typing.Any=defaultNamedOptArg, ShowBubbleSize: typing.Any=defaultNamedOptArg
			, Separator: typing.Any=defaultNamedOptArg) -> None:
		...
	def ApplyLayout(self, Layout: int=defaultNamedNotOptArg, varChartType: typing.Any=defaultNamedOptArg) -> None:
		...
	def AreaGroups(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def AutoFormat(self, rGallery: int=defaultNamedNotOptArg, varFormat: typing.Any=defaultNamedOptArg) -> None:
		...
	def Axes(self, Type: typing.Any=defaultNamedNotOptArg, AxisGroup: XlAxisGroup=1) -> Dispatch:
		...
	def BarGroups(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def ChartWizard(self, varSource: typing.Any=defaultNamedOptArg, varGallery: typing.Any=defaultNamedOptArg, varFormat: typing.Any=defaultNamedOptArg, varPlotBy: typing.Any=defaultNamedOptArg
			, varCategoryLabels: typing.Any=defaultNamedOptArg, varSeriesLabels: typing.Any=defaultNamedOptArg, varHasLegend: typing.Any=defaultNamedOptArg, varTitle: typing.Any=defaultNamedOptArg, varCategoryTitle: typing.Any=defaultNamedOptArg
			, varValueTitle: typing.Any=defaultNamedOptArg, varExtraTitle: typing.Any=defaultNamedOptArg) -> None:
		...
	def ClearToMatchColorStyle(self) -> None:
		...
	def ClearToMatchStyle(self) -> None:
		...
	def ColumnGroups(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def Copy(self) -> typing.Any:
		...
	def CopyPicture(self, Appearance: int=1, Format: int=-4147, Size: int=2) -> None:
		...
	def Delete(self) -> typing.Any:
		...
	def DeleteHiddenContent(self) -> None:
		...
	def DoughnutGroups(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def Export(self, bstr: str=defaultNamedNotOptArg, varFilterName: typing.Any=defaultNamedOptArg, varInteractive: typing.Any=defaultNamedOptArg) -> bool:
		...
	def FullSeriesCollection(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def GetChartElement(self, x: int=defaultNamedNotOptArg, y: int=defaultNamedNotOptArg, ElementID: int=defaultNamedNotOptArg, Arg1: int=defaultNamedNotOptArg
			, Arg2: int=defaultNamedNotOptArg) -> None:
		...
	# The method GetChartGroups is actually a property, but must be used as a method to correctly pass the arguments
	def GetChartGroups(self, pvarIndex: typing.Any=defaultNamedOptArg, varIgallery: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	# The method GetHasAxis is actually a property, but must be used as a method to correctly pass the arguments
	def GetHasAxis(self, axisType: typing.Any=defaultNamedOptArg, AxisGroup: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	# Result is of type IMsoWalls
	# The method GetWalls is actually a property, but must be used as a method to correctly pass the arguments
	def GetWalls(self, fBackWall: bool=True) -> IMsoWalls:
		...
	def LineGroups(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def PieGroups(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def RadarGroups(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def Refresh(self) -> None:
		...
	def RefreshPivotTable(self) -> None:
		...
	def SaveChartTemplate(self, bstrFileName: str=defaultNamedNotOptArg) -> None:
		...
	def Select(self, Replace: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def SeriesCollection(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def SetDefaultChart(self, varName: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def SetElement(self, RHS: MsoChartElementType=defaultNamedNotOptArg) -> None:
		...
	# The method SetHasAxis is actually a property, but must be used as a method to correctly pass the arguments
	def SetHasAxis(self, axisType: typing.Any=defaultNamedNotOptArg, AxisGroup: typing.Any=defaultNamedOptArg, arg2: typing.Any=defaultUnnamedArg) -> None:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def SetSourceData(self, Source: str=defaultNamedNotOptArg, PlotBy: typing.Any=defaultNamedOptArg) -> None:
		...
	def XYGroups(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def _ApplyDataLabels(self, Type: XlDataLabelsType=2, IMsoLegendKey: typing.Any=defaultNamedOptArg, AutoText: typing.Any=defaultNamedOptArg, HasLeaderLines: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
		# Method 'Area3DGroup' returns object of type 'IMsoChartGroup'
	Area3DGroup: IMsoChartGroup
	AutoScaling: typing.Any
		# Method 'BackWall' returns object of type 'IMsoWalls'
	BackWall: IMsoWalls
		# Method 'Bar3DGroup' returns object of type 'IMsoChartGroup'
	Bar3DGroup: IMsoChartGroup
		# Method 'BarShape' returns enumeration of type 'XlBarShape'
	BarShape: XlBarShape
		# Method 'CategoryLabelLevel' returns enumeration of type 'XlCategoryLabelLevel'
	CategoryLabelLevel: XlCategoryLabelLevel
		# Method 'ChartArea' returns object of type 'IMsoChartArea'
	ChartArea: IMsoChartArea
	ChartColor: typing.Any
		# Method 'ChartData' returns object of type 'IMsoChartData'
	ChartData: IMsoChartData
	ChartGroups: typing.Any
	ChartStyle: typing.Any
		# Method 'ChartTitle' returns object of type 'IMsoChartTitle'
	ChartTitle: IMsoChartTitle
		# Method 'ChartType' returns enumeration of type 'XlChartType'
	ChartType: XlChartType
		# Method 'Column3DGroup' returns object of type 'IMsoChartGroup'
	Column3DGroup: IMsoChartGroup
		# Method 'Corners' returns object of type 'IMsoCorners'
	Corners: IMsoCorners
	Creator: typing.Any
		# Method 'DataTable' returns object of type 'IMsoDataTable'
	DataTable: IMsoDataTable
	DepthPercent: typing.Any
		# Method 'DisplayBlanksAs' returns enumeration of type 'XlDisplayBlanksAs'
	DisplayBlanksAs: XlDisplayBlanksAs
	DisplayValueNotAvailableAsBlank: typing.Any
	Elevation: typing.Any
		# Method 'Floor' returns object of type 'IMsoFloor'
	Floor: IMsoFloor
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	GapDepth: typing.Any
	HasAxis: typing.Any
	HasDataTable: typing.Any
	HasHiddenContent: typing.Any
	HasLegend: typing.Any
	HasPivotFields: typing.Any
	HasTitle: typing.Any
	HeightPercent: typing.Any
		# Method 'Legend' returns object of type 'IMsoLegend'
	Legend: IMsoLegend
		# Method 'Line3DGroup' returns object of type 'IMsoChartGroup'
	Line3DGroup: IMsoChartGroup
	Parent: typing.Any
	Perspective: typing.Any
		# Method 'Pie3DGroup' returns object of type 'IMsoChartGroup'
	Pie3DGroup: IMsoChartGroup
	PivotLayout: typing.Any
		# Method 'PlotArea' returns object of type 'IMsoPlotArea'
	PlotArea: IMsoPlotArea
		# Method 'PlotBy' returns enumeration of type 'XlRowCol'
	PlotBy: XlRowCol
	PlotVisibleOnly: typing.Any
	ProtectChartObjects: typing.Any
	ProtectData: typing.Any
	ProtectFormatting: typing.Any
	ProtectGoalSeek: typing.Any
	ProtectSelection: typing.Any
	RightAngleAxes: typing.Any
	Rotation: typing.Any
	Selection: typing.Any
		# Method 'SeriesNameLevel' returns enumeration of type 'XlSeriesNameLevel'
	SeriesNameLevel: XlSeriesNameLevel
		# Method 'Shapes' returns object of type 'Shapes'
	Shapes: Shapes
	ShowAllFieldButtons: typing.Any
	ShowAxisFieldButtons: typing.Any
	ShowDataLabelsOverMaximum: typing.Any
	ShowExpandCollapseEntireFieldButtons: typing.Any
	ShowLegendFieldButtons: typing.Any
	ShowReportFilterFieldButtons: typing.Any
	ShowValueFieldButtons: typing.Any
		# Method 'SideWall' returns object of type 'IMsoWalls'
	SideWall: IMsoWalls
	SubType: typing.Any
		# Method 'SurfaceGroup' returns object of type 'IMsoChartGroup'
	SurfaceGroup: IMsoChartGroup
	Type: typing.Any
		# Method 'Walls' returns object of type 'IMsoWalls'
	Walls: IMsoWalls
	def __iter__(self):
		...

class IMsoChartArea(typing.Protocol):

	def Clear(self) -> typing.Any:
		...
	def ClearContents(self) -> typing.Any:
		...
	def ClearFormats(self) -> typing.Any:
		...
	def Copy(self) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	Application: typing.Any
	AutoScaleFont: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Font' returns object of type 'ChartFont'
	Font: ChartFont
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Height: typing.Any
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Left: typing.Any
	Name: typing.Any
	Parent: typing.Any
	RoundedCorners: typing.Any
	Shadow: typing.Any
	Top: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class IMsoChartData(typing.Protocol):

	def Activate(self) -> None:
		...
	def ActivateChartDataWindow(self) -> None:
		...
	def BreakLink(self) -> None:
		...
	IsLinked: typing.Any
	Workbook: typing.Any
	def __iter__(self):
		...

class IMsoChartFormat(typing.Protocol):

		# Method 'Adjustments' returns object of type 'Adjustments'
	Adjustments: Adjustments
	Application: typing.Any
		# Method 'AutoShapeType' returns enumeration of type 'MsoAutoShapeType'
	AutoShapeType: MsoAutoShapeType
	Creator: typing.Any
		# Method 'Fill' returns object of type 'FillFormat'
	Fill: FillFormat
		# Method 'Glow' returns object of type 'GlowFormat'
	Glow: GlowFormat
		# Method 'Line' returns object of type 'LineFormat'
	Line: LineFormat
	Parent: typing.Any
		# Method 'PictureFormat' returns object of type 'PictureFormat'
	PictureFormat: PictureFormat
		# Method 'Shadow' returns object of type 'ShadowFormat'
	Shadow: ShadowFormat
		# Method 'SoftEdge' returns object of type 'SoftEdgeFormat'
	SoftEdge: SoftEdgeFormat
		# Method 'TextFrame2' returns object of type 'TextFrame2'
	TextFrame2: TextFrame2
		# Method 'ThreeD' returns object of type 'ThreeDFormat'
	ThreeD: ThreeDFormat
	def __iter__(self):
		...

class IMsoChartGroup(typing.Protocol):

	def CategoryCollection(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def FullCategoryCollection(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def SeriesCollection(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	Application: typing.Any
	AxisGroup: typing.Any
	BinWidthValue: typing.Any
	BinsCountValue: typing.Any
	BinsOverflowEnabled: typing.Any
	BinsOverflowValue: typing.Any
		# Method 'BinsType' returns enumeration of type 'XlBinsType'
	BinsType: XlBinsType
	BinsUnderflowEnabled: typing.Any
	BinsUnderflowValue: typing.Any
	BubbleScale: typing.Any
	Creator: typing.Any
	DoughnutHoleSize: typing.Any
		# Method 'DownBars' returns object of type 'IMsoDownBars'
	DownBars: IMsoDownBars
		# Method 'DropLines' returns object of type 'IMsoDropLines'
	DropLines: IMsoDropLines
	FirstSliceAngle: typing.Any
	GapWidth: typing.Any
	Has3DShading: typing.Any
	HasDropLines: typing.Any
	HasHiLoLines: typing.Any
	HasRadarAxisLabels: typing.Any
	HasSeriesLines: typing.Any
	HasUpDownBars: typing.Any
		# Method 'HiLoLines' returns object of type 'IMsoHiLoLines'
	HiLoLines: IMsoHiLoLines
	Index: typing.Any
	Overlap: typing.Any
	Parent: typing.Any
	RadarAxisLabels: typing.Any
	SecondPlotSize: typing.Any
		# Method 'SeriesLines' returns object of type 'IMsoSeriesLines'
	SeriesLines: IMsoSeriesLines
	ShowNegativeBubbles: typing.Any
		# Method 'SizeRepresents' returns enumeration of type 'XlSizeRepresents'
	SizeRepresents: XlSizeRepresents
		# Method 'SplitType' returns enumeration of type 'XlChartSplitType'
	SplitType: XlChartSplitType
	SplitValue: typing.Any
	SubType: typing.Any
	Type: typing.Any
		# Method 'UpBars' returns object of type 'IMsoUpBars'
	UpBars: IMsoUpBars
	VaryByCategories: typing.Any
	def __iter__(self):
		...

class IMsoChartTitle(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	# Result is of type IMsoCharacters
	# The method GetCharacters is actually a property, but must be used as a method to correctly pass the arguments
	def GetCharacters(self, Start: typing.Any=defaultNamedOptArg, Length: typing.Any=defaultNamedOptArg) -> IMsoCharacters:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	AutoScaleFont: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Caption: typing.Any
		# Method 'Characters' returns object of type 'IMsoCharacters'
	Characters: IMsoCharacters
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Font' returns object of type 'ChartFont'
	Font: ChartFont
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Formula: typing.Any
	FormulaLocal: typing.Any
	FormulaR1C1: typing.Any
	FormulaR1C1Local: typing.Any
	Height: typing.Any
	HorizontalAlignment: typing.Any
	IncludeInLayout: typing.Any
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Left: typing.Any
	Name: typing.Any
	Orientation: typing.Any
	Parent: typing.Any
		# Method 'Position' returns enumeration of type 'XlChartElementPosition'
	Position: XlChartElementPosition
	ReadingOrder: typing.Any
	Shadow: typing.Any
	Text: typing.Any
	Top: typing.Any
	VerticalAlignment: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class IMsoContactCard(typing.Protocol):

	Address: typing.Any
		# Method 'AddressType' returns enumeration of type 'MsoContactCardAddressType'
	AddressType: MsoContactCardAddressType
	Application: typing.Any
		# Method 'CardType' returns enumeration of type 'MsoContactCardType'
	CardType: MsoContactCardType
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoCorners(typing.Protocol):

	def Select(self) -> typing.Any:
		...
	Application: typing.Any
	Creator: typing.Any
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoDataLabel(typing.Protocol):

	Application: typing.Any
	AutoScaleFont: typing.Any
	AutoText: typing.Any
	Border: typing.Any
	Caption: typing.Any
	Characters: typing.Any
	Creator: typing.Any
	Fill: typing.Any
	Font: typing.Any
	Format: typing.Any
	Formula: typing.Any
	FormulaLocal: typing.Any
	FormulaR1C1: typing.Any
	FormulaR1C1Local: typing.Any
	Height: typing.Any
	HorizontalAlignment: typing.Any
	Interior: typing.Any
	Left: typing.Any
	Name: typing.Any
	NumberFormat: typing.Any
	NumberFormatLinked: typing.Any
	NumberFormatLocal: typing.Any
	Orientation: typing.Any
	Parent: typing.Any
	Position: typing.Any
	ReadingOrder: typing.Any
	Separator: typing.Any
	Shadow: typing.Any
	ShowBubbleSize: typing.Any
	ShowCategoryName: typing.Any
	ShowLegendKey: typing.Any
	ShowPercentage: typing.Any
	ShowRange: typing.Any
	ShowSeriesName: typing.Any
	ShowValue: typing.Any
	Text: typing.Any
	Top: typing.Any
	Type: typing.Any
	VerticalAlignment: typing.Any
	Width: typing.Any
	_Height: typing.Any
	_Width: typing.Any
	def __iter__(self):
		...

class IMsoDataLabels(typing.Protocol):

	Application: typing.Any
	AutoScaleFont: typing.Any
	AutoText: typing.Any
	Border: typing.Any
	Characters: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Fill: typing.Any
	Font: typing.Any
	Format: typing.Any
	HorizontalAlignment: typing.Any
	Interior: typing.Any
	Name: typing.Any
	NumberFormat: typing.Any
	NumberFormatLinked: typing.Any
	NumberFormatLocal: typing.Any
	Orientation: typing.Any
	Parent: typing.Any
	Position: typing.Any
	ReadingOrder: typing.Any
	Separator: typing.Any
	Shadow: typing.Any
	ShowBubbleSize: typing.Any
	ShowCategoryName: typing.Any
	ShowLegendKey: typing.Any
	ShowPercentage: typing.Any
	ShowRange: typing.Any
	ShowSeriesName: typing.Any
	ShowValue: typing.Any
	Type: typing.Any
	VerticalAlignment: typing.Any
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class IMsoDataTable(typing.Protocol):

	def Delete(self) -> None:
		...
	def Select(self) -> None:
		...
	Application: typing.Any
	AutoScaleFont: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Font' returns object of type 'ChartFont'
	Font: ChartFont
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	HasBorderHorizontal: typing.Any
	HasBorderOutline: typing.Any
	HasBorderVertical: typing.Any
	Parent: typing.Any
	ShowLegendKey: typing.Any
	def __iter__(self):
		...

class IMsoDiagram(typing.Protocol):

	def Convert(self, Type: MsoDiagramType=defaultNamedNotOptArg) -> None:
		...
	def FitText(self) -> None:
		...
	Application: typing.Any
		# Method 'AutoFormat' returns enumeration of type 'MsoTriState'
	AutoFormat: MsoTriState
		# Method 'AutoLayout' returns enumeration of type 'MsoTriState'
	AutoLayout: MsoTriState
	Creator: typing.Any
		# Method 'Nodes' returns object of type 'DiagramNodes'
	Nodes: DiagramNodes
	Parent: typing.Any
		# Method 'Reverse' returns enumeration of type 'MsoTriState'
	Reverse: MsoTriState
		# Method 'Type' returns enumeration of type 'MsoDiagramType'
	Type: MsoDiagramType
	def __iter__(self):
		...

class IMsoDispCagNotifySink(typing.Protocol):

	def InsertClip(self, pClipMoniker: typing.Any=defaultNamedNotOptArg, pItemMoniker: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def WindowIsClosing(self) -> None:
		...
	def __iter__(self):
		...

class IMsoDisplayUnitLabel(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	# Result is of type IMsoCharacters
	# The method GetCharacters is actually a property, but must be used as a method to correctly pass the arguments
	def GetCharacters(self, Start: typing.Any=defaultNamedOptArg, Length: typing.Any=defaultNamedOptArg) -> IMsoCharacters:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	AutoScaleFont: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Caption: typing.Any
		# Method 'Characters' returns object of type 'IMsoCharacters'
	Characters: IMsoCharacters
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Font' returns object of type 'ChartFont'
	Font: ChartFont
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Formula: typing.Any
	FormulaLocal: typing.Any
	FormulaR1C1: typing.Any
	FormulaR1C1Local: typing.Any
	Height: typing.Any
	HorizontalAlignment: typing.Any
	IncludeInLayout: typing.Any
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Left: typing.Any
	Name: typing.Any
	Orientation: typing.Any
	Parent: typing.Any
		# Method 'Position' returns enumeration of type 'XlChartElementPosition'
	Position: XlChartElementPosition
	ReadingOrder: typing.Any
	Shadow: typing.Any
	Text: typing.Any
	Top: typing.Any
	VerticalAlignment: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class IMsoDownBars(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoDropLines(typing.Protocol):

	def Delete(self) -> None:
		...
	def Select(self) -> None:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoEServicesDialog(typing.Protocol):

	def AddTrustedDomain(self, Domain: str=defaultNamedNotOptArg) -> None:
		...
	def Close(self, ApplyWebComponentChanges: bool=False) -> None:
		...
	Application: typing.Any
	ApplicationName: typing.Any
	ClipArt: typing.Any
	WebComponent: typing.Any
	def __iter__(self):
		...

class IMsoEnvelopeVB(typing.Protocol):

	CommandBars: typing.Any
	Introduction: typing.Any
	Item: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		...

class IMsoEnvelopeVBEvents:

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnEnvelopeShow(self):
#	def OnEnvelopeHide(self):
	...


class IMsoErrorBars(typing.Protocol):

	Application: typing.Any
	Border: typing.Any
	Creator: typing.Any
	EndStyle: typing.Any
	Format: typing.Any
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoFloor(typing.Protocol):

	def ClearFormats(self) -> typing.Any:
		...
	def Paste(self) -> None:
		...
	def Select(self) -> typing.Any:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Name: typing.Any
	Parent: typing.Any
	PictureType: typing.Any
	Thickness: typing.Any
	def __iter__(self):
		...

class IMsoHiLoLines(typing.Protocol):

	def Delete(self) -> None:
		...
	def Select(self) -> None:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoHyperlinks(typing.Protocol):

	def __iter__(self):
		...

class IMsoInterior(typing.Protocol):

	Application: typing.Any
	Color: typing.Any
	ColorIndex: typing.Any
	Creator: typing.Any
	InvertIfNegative: typing.Any
	Parent: typing.Any
	Pattern: typing.Any
	PatternColor: typing.Any
	PatternColorIndex: typing.Any
	def __iter__(self):
		...

class IMsoLeaderLines(typing.Protocol):

	def Delete(self) -> None:
		...
	def Select(self) -> None:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoLegend(typing.Protocol):

	def Clear(self) -> typing.Any:
		...
	def Delete(self) -> typing.Any:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def LegendEntries(self, Index: typing.Any=defaultNamedOptArg) -> Dispatch:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	AutoScaleFont: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Font' returns object of type 'ChartFont'
	Font: ChartFont
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Height: typing.Any
	IncludeInLayout: typing.Any
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Left: typing.Any
	Name: typing.Any
	Parent: typing.Any
		# Method 'Position' returns enumeration of type 'XlLegendPosition'
	Position: XlLegendPosition
	Shadow: typing.Any
	Top: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class IMsoLegendKey(typing.Protocol):

	Application: typing.Any
	Border: typing.Any
	Creator: typing.Any
	Fill: typing.Any
	Format: typing.Any
	Height: typing.Any
	Interior: typing.Any
	InvertIfNegative: typing.Any
	Left: typing.Any
	MarkerBackgroundColor: typing.Any
	MarkerBackgroundColorIndex: typing.Any
	MarkerForegroundColor: typing.Any
	MarkerForegroundColorIndex: typing.Any
	MarkerSize: typing.Any
	MarkerStyle: typing.Any
	Parent: typing.Any
	PictureType: typing.Any
	PictureUnit: typing.Any
	PictureUnit2: typing.Any
	Shadow: typing.Any
	Smooth: typing.Any
	Top: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class IMsoPlotArea(typing.Protocol):

	def ClearFormats(self) -> typing.Any:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Height: typing.Any
	InsideHeight: typing.Any
	InsideLeft: typing.Any
	InsideTop: typing.Any
	InsideWidth: typing.Any
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Left: typing.Any
	Name: typing.Any
	Parent: typing.Any
		# Method 'Position' returns enumeration of type 'XlChartElementPosition'
	Position: XlChartElementPosition
	Top: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class IMsoSeries(typing.Protocol):

	Application: typing.Any
	ApplyPictToEnd: typing.Any
	ApplyPictToFront: typing.Any
	ApplyPictToSides: typing.Any
	AxisGroup: typing.Any
	BarShape: typing.Any
	Border: typing.Any
	BubbleSizes: typing.Any
	ChartType: typing.Any
	Creator: typing.Any
	ErrorBars: typing.Any
	Explosion: typing.Any
	Fill: typing.Any
	Format: typing.Any
	Formula: typing.Any
	FormulaLocal: typing.Any
	FormulaR1C1: typing.Any
	FormulaR1C1Local: typing.Any
	GeoMappingLevel: typing.Any
	GeoProjectionType: typing.Any
	Has3DEffect: typing.Any
	HasDataLabels: typing.Any
	HasErrorBars: typing.Any
	HasLeaderLines: typing.Any
	Interior: typing.Any
	InvertColor: typing.Any
	InvertColorIndex: typing.Any
	InvertIfNegative: typing.Any
	IsFiltered: typing.Any
	LeaderLines: typing.Any
	MarkerBackgroundColor: typing.Any
	MarkerBackgroundColorIndex: typing.Any
	MarkerForegroundColor: typing.Any
	MarkerForegroundColorIndex: typing.Any
	MarkerSize: typing.Any
	MarkerStyle: typing.Any
	Name: typing.Any
	Parent: typing.Any
	ParentDataLabelOption: typing.Any
	PictureType: typing.Any
	PictureUnit: typing.Any
	PictureUnit2: typing.Any
	PlotColorIndex: typing.Any
	PlotOrder: typing.Any
	QuartileCalculationInclusiveMedian: typing.Any
	RegionLabelOption: typing.Any
	SeriesColorGradientStyle: typing.Any
	SeriesColorMaxGradientStop: typing.Any
	SeriesColorMidGradientStop: typing.Any
	SeriesColorMinGradientStop: typing.Any
	Shadow: typing.Any
	Smooth: typing.Any
	Type: typing.Any
	ValueSortOrder: typing.Any
	Values: typing.Any
	XValues: typing.Any
	def __iter__(self):
		...

class IMsoSeriesLines(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoTickLabels(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	Alignment: typing.Any
	Application: typing.Any
	AutoScaleFont: typing.Any
	Creator: typing.Any
	Depth: typing.Any
		# Method 'Font' returns object of type 'ChartFont'
	Font: ChartFont
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
	MultiLevel: typing.Any
	Name: typing.Any
	NumberFormat: typing.Any
	NumberFormatLinked: typing.Any
	NumberFormatLocal: typing.Any
	Offset: typing.Any
		# Method 'Orientation' returns enumeration of type 'XlTickLabelOrientation'
	Orientation: XlTickLabelOrientation
	Parent: typing.Any
	ReadingOrder: typing.Any
	def __iter__(self):
		...

class IMsoTrendline(typing.Protocol):

	Application: typing.Any
	Backward: typing.Any
	Backward2: typing.Any
	Border: typing.Any
	Creator: typing.Any
	DataLabel: typing.Any
	DisplayEquation: typing.Any
	DisplayRSquared: typing.Any
	Format: typing.Any
	Forward: typing.Any
	Forward2: typing.Any
	Index: typing.Any
	Intercept: typing.Any
	InterceptIsAuto: typing.Any
	Name: typing.Any
	NameIsAuto: typing.Any
	Order: typing.Any
	Parent: typing.Any
	Period: typing.Any
	Type: typing.Any
	def __iter__(self):
		...

class IMsoUpBars(typing.Protocol):

	def Delete(self) -> typing.Any:
		...
	def GetProperty(self, bstrId: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Select(self) -> typing.Any:
		...
	def SetProperty(self, bstrId: str=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class IMsoWalls(typing.Protocol):

	def ClearFormats(self) -> typing.Any:
		...
	def Paste(self) -> None:
		...
	def Select(self) -> typing.Any:
		...
	Application: typing.Any
		# Method 'Border' returns object of type 'IMsoBorder'
	Border: IMsoBorder
	Creator: typing.Any
		# Method 'Fill' returns object of type 'ChartFillFormat'
	Fill: ChartFillFormat
		# Method 'Format' returns object of type 'IMsoChartFormat'
	Format: IMsoChartFormat
		# Method 'Interior' returns object of type 'IMsoInterior'
	Interior: IMsoInterior
	Name: typing.Any
	Parent: typing.Any
	PictureType: typing.Any
	PictureUnit: typing.Any
	Thickness: typing.Any
	def __iter__(self):
		...

class IRibbonControl(typing.Protocol):

	Context: typing.Any
	Id: typing.Any
	Tag: typing.Any
	def __iter__(self):
		...

class IRibbonExtensibility(typing.Protocol):

	def GetCustomUI(self, RibbonID: str=defaultNamedNotOptArg) -> str:
		...
	def __iter__(self):
		...

class IRibbonUI(typing.Protocol):

	def ActivateTab(self, ControlID: str=defaultNamedNotOptArg) -> None:
		...
	def ActivateTabMso(self, ControlID: str=defaultNamedNotOptArg) -> None:
		...
	def ActivateTabQ(self, ControlID: str=defaultNamedNotOptArg, Namespace: str=defaultNamedNotOptArg) -> None:
		...
	def Invalidate(self) -> None:
		...
	def InvalidateControl(self, ControlID: str=defaultNamedNotOptArg) -> None:
		...
	def InvalidateControlMso(self, ControlID: str=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class ISensitivityLabel(typing.Protocol):

	# Result is of type LabelInfo
	def CreateLabelInfo(self) -> LabelInfo:
		...
	# Result is of type LabelInfo
	def GetLabel(self) -> LabelInfo:
		...
	def SetLabel(self, LabelInfo: LabelInfo=defaultNamedNotOptArg, Context: Dispatch=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
		# Method 'SensitivityLabelError' returns enumeration of type 'MsoSensitivityLabelError'
	SensitivityLabelError: MsoSensitivityLabelError
	def __iter__(self):
		...

class ISensitivityLabelEvents(typing.Protocol):

	def LabelChanged(self, OldLabelInfo: LabelInfo=defaultNamedNotOptArg, NewLabelInfo: LabelInfo=defaultNamedNotOptArg, HResult: int=defaultNamedNotOptArg, Context: Dispatch=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class LabelInfo(typing.Protocol):

	ActionId: typing.Any
	Application: typing.Any
		# Method 'AssignmentMethod' returns enumeration of type 'MsoAssignmentMethod'
	AssignmentMethod: MsoAssignmentMethod
	ContentBits: typing.Any
	Creator: typing.Any
	IsEnabled: typing.Any
	Justification: typing.Any
	LabelId: typing.Any
	LabelName: typing.Any
	SetDate: typing.Any
	SiteId: typing.Any
	# Default property for this class is 'LabelId'
	def __call__(self):
		...
	def __iter__(self):
		...

class LanguageSettings(typing.Protocol):

	# The method LanguageID is actually a property, but must be used as a method to correctly pass the arguments
	def LanguageID(self, Id: MsoAppLanguageID=defaultNamedNotOptArg) -> int:
		...
	# The method LanguagePreferredForEditing is actually a property, but must be used as a method to correctly pass the arguments
	def LanguagePreferredForEditing(self, lid: MsoLanguageID=defaultNamedNotOptArg) -> bool:
		...
	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class LegendEntries(typing.Protocol):

	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class LegendEntry(typing.Protocol):

	Application: typing.Any
	AutoScaleFont: typing.Any
	Creator: typing.Any
	Font: typing.Any
	Format: typing.Any
	Height: typing.Any
	Index: typing.Any
	Left: typing.Any
	LegendKey: typing.Any
	Parent: typing.Any
	Top: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class LineFormat(typing.Protocol):

	Application: typing.Any
		# Method 'BackColor' returns object of type 'ColorFormat'
	BackColor: ColorFormat
		# Method 'BeginArrowheadLength' returns enumeration of type 'MsoArrowheadLength'
	BeginArrowheadLength: MsoArrowheadLength
		# Method 'BeginArrowheadStyle' returns enumeration of type 'MsoArrowheadStyle'
	BeginArrowheadStyle: MsoArrowheadStyle
		# Method 'BeginArrowheadWidth' returns enumeration of type 'MsoArrowheadWidth'
	BeginArrowheadWidth: MsoArrowheadWidth
	Creator: typing.Any
		# Method 'DashStyle' returns enumeration of type 'MsoLineDashStyle'
	DashStyle: MsoLineDashStyle
		# Method 'EndArrowheadLength' returns enumeration of type 'MsoArrowheadLength'
	EndArrowheadLength: MsoArrowheadLength
		# Method 'EndArrowheadStyle' returns enumeration of type 'MsoArrowheadStyle'
	EndArrowheadStyle: MsoArrowheadStyle
		# Method 'EndArrowheadWidth' returns enumeration of type 'MsoArrowheadWidth'
	EndArrowheadWidth: MsoArrowheadWidth
		# Method 'ForeColor' returns object of type 'ColorFormat'
	ForeColor: ColorFormat
		# Method 'InsetPen' returns enumeration of type 'MsoTriState'
	InsetPen: MsoTriState
	Parent: typing.Any
		# Method 'Pattern' returns enumeration of type 'MsoPatternType'
	Pattern: MsoPatternType
		# Method 'Style' returns enumeration of type 'MsoLineStyle'
	Style: MsoLineStyle
	Transparency: typing.Any
		# Method 'Visible' returns enumeration of type 'MsoTriState'
	Visible: MsoTriState
	Weight: typing.Any
	def __iter__(self):
		...

class MetaProperties(typing.Protocol):

	# Result is of type MetaProperty
	def GetItemByInternalName(self, InternalName: str=defaultNamedNotOptArg) -> MetaProperty:
		...
	# Result is of type MetaProperty
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> MetaProperty:
		...
	def Validate(self) -> str:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	SchemaXml: typing.Any
	ValidationError: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> MetaProperty:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class MetaProperty(typing.Protocol):

	def Validate(self) -> str:
		...
	Application: typing.Any
	Creator: typing.Any
	Id: typing.Any
	IsReadOnly: typing.Any
	IsRequired: typing.Any
	Name: typing.Any
	Parent: typing.Any
		# Method 'Type' returns enumeration of type 'MsoMetaPropertyType'
	Type: MsoMetaPropertyType
	ValidationError: typing.Any
	Value: typing.Any
	# Default property for this class is 'Value'
	def __call__(self):
		...
	def __iter__(self):
		...

class Model3DFormat(typing.Protocol):

	def IncrementRotationX(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementRotationY(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementRotationZ(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def ResetModel(self, ResetSize: bool=False) -> None:
		...
	Application: typing.Any
		# Method 'AutoFit' returns enumeration of type 'MsoTriState'
	AutoFit: MsoTriState
	CameraPositionX: typing.Any
	CameraPositionY: typing.Any
	CameraPositionZ: typing.Any
	Creator: typing.Any
	FieldOfView: typing.Any
	LookAtPointX: typing.Any
	LookAtPointY: typing.Any
	LookAtPointZ: typing.Any
	Parent: typing.Any
	RotationX: typing.Any
	RotationY: typing.Any
	RotationZ: typing.Any
	def __iter__(self):
		...

class MsoDebugOptions(typing.Protocol):

	def AddIgnoredAssertTag(self, bstrTagToIgnore: str=defaultNamedNotOptArg) -> None:
		...
	def RemoveIgnoredAssertTag(self, bstrTagToIgnore: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	FeatureReports: typing.Any
	OutputToDebugger: typing.Any
	OutputToFile: typing.Any
	OutputToMessageBox: typing.Any
	UnitTestManager: typing.Any
	def __iter__(self):
		...

class MsoDebugOptions_UT(typing.Protocol):

	# Result is of type MsoDebugOptions_UTRunResult
	def Run(self) -> MsoDebugOptions_UTRunResult:
		...
	Application: typing.Any
	CollectionName: typing.Any
	Creator: typing.Any
	Name: typing.Any
	# Default property for this class is 'Name'
	def __call__(self):
		...
	def __iter__(self):
		...

class MsoDebugOptions_UTManager(typing.Protocol):

	def NotifyEndOfTestSuiteRun(self) -> None:
		...
	def NotifyStartOfTestSuiteRun(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	ReportErrors: typing.Any
		# Method 'UnitTests' returns object of type 'MsoDebugOptions_UTs'
	UnitTests: MsoDebugOptions_UTs
	# Default property for this class is 'UnitTests'
	def __call__(self):
		...
	def __iter__(self):
		...

class MsoDebugOptions_UTRunResult(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	ErrorString: typing.Any
	Passed: typing.Any
	# Default property for this class is 'Passed'
	def __call__(self):
		...
	def __iter__(self):
		...

class MsoDebugOptions_UTs(typing.Protocol):

	# Result is of type MsoDebugOptions_UTs
	def GetMatchingUnitTestsInCollection(self, bstrCollectionName: str=defaultNamedNotOptArg, bstrUnitTestNameFilter: str=defaultNamedNotOptArg) -> MsoDebugOptions_UTs:
		...
	# Result is of type MsoDebugOptions_UT
	def GetUnitTest(self, bstrCollectionName: str=defaultNamedNotOptArg, bstrUnitTestName: str=defaultNamedNotOptArg) -> MsoDebugOptions_UT:
		...
	# Result is of type MsoDebugOptions_UTs
	def GetUnitTestsInCollection(self, bstrCollectionName: str=defaultNamedNotOptArg) -> MsoDebugOptions_UTs:
		...
	# Result is of type MsoDebugOptions_UT
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> MsoDebugOptions_UT:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> MsoDebugOptions_UT:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class NewFile(typing.Protocol):

	def Add(self, FileName: str=defaultNamedNotOptArg, Section: typing.Any=defaultNamedOptArg, DisplayName: typing.Any=defaultNamedOptArg, Action: typing.Any=defaultNamedOptArg) -> bool:
		...
	def Remove(self, FileName: str=defaultNamedNotOptArg, Section: typing.Any=defaultNamedOptArg, DisplayName: typing.Any=defaultNamedOptArg, Action: typing.Any=defaultNamedOptArg) -> bool:
		...
	Application: typing.Any
	Creator: typing.Any
	def __iter__(self):
		...

class ODSOColumn(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Index: typing.Any
	Name: typing.Any
	Parent: typing.Any
	Value: typing.Any
	# Default property for this class is 'Value'
	def __call__(self):
		...
	def __iter__(self):
		...

class ODSOColumns(typing.Protocol):

	def Item(self, varIndex: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ODSOFilter(typing.Protocol):

	Application: typing.Any
	Column: typing.Any
	CompareTo: typing.Any
		# Method 'Comparison' returns enumeration of type 'MsoFilterComparison'
	Comparison: MsoFilterComparison
		# Method 'Conjunction' returns enumeration of type 'MsoFilterConjunction'
	Conjunction: MsoFilterConjunction
	Creator: typing.Any
	Index: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class ODSOFilters(typing.Protocol):

	def Add(self, Column: str=defaultNamedNotOptArg, Comparison: MsoFilterComparison=defaultNamedNotOptArg, Conjunction: MsoFilterConjunction=defaultNamedNotOptArg, bstrCompareTo: str=''
			, DeferUpdate: bool=False) -> None:
		...
	def Delete(self, Index: int=defaultNamedNotOptArg, DeferUpdate: bool=False) -> None:
		...
	def Item(self, Index: int=defaultNamedNotOptArg) -> Dispatch:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class OfficeDataSourceObject(typing.Protocol):

	def ApplyFilter(self) -> None:
		...
	def Move(self, MsoMoveRow: MsoMoveRow=defaultNamedNotOptArg, RowNbr: int=1) -> int:
		...
	def Open(self, bstrSrc: str='', bstrConnect: str='', bstrTable: str='', fOpenExclusive: int=0
			, fNeverPrompt: int=1) -> None:
		...
	def SetSortOrder(self, SortField1: str=defaultNamedNotOptArg, SortAscending1: bool=True, SortField2: str='', SortAscending2: bool=True
			, SortField3: str='', SortAscending3: bool=True) -> None:
		...
	Columns: typing.Any
	ConnectString: typing.Any
	DataSource: typing.Any
	Filters: typing.Any
	RowCount: typing.Any
	Table: typing.Any
	def __iter__(self):
		...

class OfficeTheme(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
		# Method 'ThemeColorScheme' returns object of type 'ThemeColorScheme'
	ThemeColorScheme: ThemeColorScheme
		# Method 'ThemeEffectScheme' returns object of type 'ThemeEffectScheme'
	ThemeEffectScheme: ThemeEffectScheme
		# Method 'ThemeFontScheme' returns object of type 'ThemeFontScheme'
	ThemeFontScheme: ThemeFontScheme
	def __iter__(self):
		...

class ParagraphFormat2(typing.Protocol):

		# Method 'Alignment' returns enumeration of type 'MsoParagraphAlignment'
	Alignment: MsoParagraphAlignment
	Application: typing.Any
		# Method 'BaselineAlignment' returns enumeration of type 'MsoBaselineAlignment'
	BaselineAlignment: MsoBaselineAlignment
		# Method 'Bullet' returns object of type 'BulletFormat2'
	Bullet: BulletFormat2
	Creator: typing.Any
		# Method 'FarEastLineBreakLevel' returns enumeration of type 'MsoTriState'
	FarEastLineBreakLevel: MsoTriState
	FirstLineIndent: typing.Any
		# Method 'HangingPunctuation' returns enumeration of type 'MsoTriState'
	HangingPunctuation: MsoTriState
	IndentLevel: typing.Any
	LeftIndent: typing.Any
		# Method 'LineRuleAfter' returns enumeration of type 'MsoTriState'
	LineRuleAfter: MsoTriState
		# Method 'LineRuleBefore' returns enumeration of type 'MsoTriState'
	LineRuleBefore: MsoTriState
		# Method 'LineRuleWithin' returns enumeration of type 'MsoTriState'
	LineRuleWithin: MsoTriState
	Parent: typing.Any
	RightIndent: typing.Any
	SpaceAfter: typing.Any
	SpaceBefore: typing.Any
	SpaceWithin: typing.Any
		# Method 'TabStops' returns object of type 'TabStops2'
	TabStops: TabStops2
		# Method 'TextDirection' returns enumeration of type 'MsoTextDirection'
	TextDirection: MsoTextDirection
		# Method 'WordWrap' returns enumeration of type 'MsoTriState'
	WordWrap: MsoTriState
	def __iter__(self):
		...

class Permission(typing.Protocol):

	# Result is of type UserPermission
	def Add(self, UserId: str=defaultNamedNotOptArg, Permission: typing.Any=defaultNamedOptArg, ExpirationDate: typing.Any=defaultNamedOptArg) -> UserPermission:
		...
	def ApplyPolicy(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	# Result is of type UserPermission
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> UserPermission:
		...
	def RemoveAll(self) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	DocumentAuthor: typing.Any
	DoubleKeyEncryptionUrl: typing.Any
	EnableTrustedBrowser: typing.Any
	Enabled: typing.Any
	Parent: typing.Any
	PermissionFromPolicy: typing.Any
	PolicyDescription: typing.Any
	PolicyName: typing.Any
	RequestPermissionURL: typing.Any
	StoreLicenses: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> UserPermission:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class PickerDialog(typing.Protocol):

	# Result is of type PickerResults
	def CreatePickerResults(self) -> PickerResults:
		...
	# Result is of type PickerResults
	def Resolve(self, TokenText: str=defaultNamedNotOptArg, duplicateDlgMode: int=defaultNamedNotOptArg) -> PickerResults:
		...
	# Result is of type PickerResults
	def Show(self, IsMultiSelect: bool=True, ExistingResults: PickerResults=0) -> PickerResults:
		...
	Application: typing.Any
	Creator: typing.Any
	DataHandlerId: typing.Any
		# Method 'Properties' returns object of type 'PickerProperties'
	Properties: PickerProperties
	Title: typing.Any
	def __iter__(self):
		...

class PickerField(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	IsHidden: typing.Any
	Name: typing.Any
		# Method 'Type' returns enumeration of type 'MsoPickerField'
	Type: MsoPickerField
	def __iter__(self):
		...

class PickerFields(typing.Protocol):

	# Result is of type PickerField
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> PickerField:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> PickerField:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class PickerProperties(typing.Protocol):

	# Result is of type PickerProperty
	def Add(self, Id: str=defaultNamedNotOptArg, Value: str=defaultNamedNotOptArg, Type: MsoPickerField=defaultNamedNotOptArg) -> PickerProperty:
		...
	# Result is of type PickerProperty
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> PickerProperty:
		...
	def Remove(self, Id: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> PickerProperty:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class PickerProperty(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Id: typing.Any
		# Method 'Type' returns enumeration of type 'MsoPickerField'
	Type: MsoPickerField
	Value: typing.Any
	# Default property for this class is 'Value'
	def __call__(self):
		...
	def __iter__(self):
		...

class PickerResult(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	DisplayName: typing.Any
	DuplicateResults: typing.Any
		# Method 'Fields' returns object of type 'PickerFields'
	Fields: PickerFields
	Id: typing.Any
	ItemData: typing.Any
	SIPId: typing.Any
	SubItems: typing.Any
	Type: typing.Any
	def __iter__(self):
		...

class PickerResults(typing.Protocol):

	# Result is of type PickerResult
	def Add(self, Id: str=defaultNamedNotOptArg, DisplayName: str=defaultNamedNotOptArg, Type: str=defaultNamedNotOptArg, SIPId: str=''
			, ItemData: typing.Any=defaultNamedOptArg, SubItems: typing.Any=defaultNamedOptArg) -> PickerResult:
		...
	# Result is of type PickerResult
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> PickerResult:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> PickerResult:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class PictureEffect(typing.Protocol):

	def Delete(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
		# Method 'EffectParameters' returns object of type 'EffectParameters'
	EffectParameters: EffectParameters
	Position: typing.Any
		# Method 'Type' returns enumeration of type 'MsoPictureEffectType'
	Type: MsoPictureEffectType
		# Method 'Visible' returns enumeration of type 'MsoTriState'
	Visible: MsoTriState
	# Default property for this class is 'Type'
	def __call__(self):
		...
	def __iter__(self):
		...

class PictureEffects(typing.Protocol):

	def Delete(self, Index: int=-1) -> None:
		...
	# Result is of type PictureEffect
	def Insert(self, EffectType: MsoPictureEffectType=defaultNamedNotOptArg, Position: int=-1) -> PictureEffect:
		...
	# Result is of type PictureEffect
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> PictureEffect:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> PictureEffect:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class PictureFormat(typing.Protocol):

	def IncrementBrightness(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementContrast(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Brightness: typing.Any
		# Method 'ColorType' returns enumeration of type 'MsoPictureColorType'
	ColorType: MsoPictureColorType
	Contrast: typing.Any
	Creator: typing.Any
		# Method 'Crop' returns object of type 'Crop'
	Crop: Crop
	CropBottom: typing.Any
	CropLeft: typing.Any
	CropRight: typing.Any
	CropTop: typing.Any
	Parent: typing.Any
	TransparencyColor: typing.Any
		# Method 'TransparentBackground' returns enumeration of type 'MsoTriState'
	TransparentBackground: MsoTriState
	def __iter__(self):
		...

class Points(typing.Protocol):

	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class PolicyItem(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Data: typing.Any
	Description: typing.Any
	Id: typing.Any
	Name: typing.Any
	Parent: typing.Any
	# Default property for this class is 'Id'
	def __call__(self):
		...
	def __iter__(self):
		...

class PropertyTest(typing.Protocol):

	Application: typing.Any
		# Method 'Condition' returns enumeration of type 'MsoCondition'
	Condition: MsoCondition
		# Method 'Connector' returns enumeration of type 'MsoConnector'
	Connector: MsoConnector
	Creator: typing.Any
	Name: typing.Any
	SecondValue: typing.Any
	Value: typing.Any
	# Default property for this class is 'Name'
	def __call__(self):
		...
	def __iter__(self):
		...

class PropertyTests(typing.Protocol):

	def Add(self, Name: str=defaultNamedNotOptArg, Condition: MsoCondition=defaultNamedNotOptArg, Value: typing.Any=defaultNamedNotOptArg, SecondValue: typing.Any=defaultNamedNotOptArg
			, Connector: MsoConnector=1) -> None:
		...
	# Result is of type PropertyTest
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> PropertyTest:
		...
	def Remove(self, Index: int=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> PropertyTest:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ReflectionFormat(typing.Protocol):

	Application: typing.Any
	Blur: typing.Any
	Creator: typing.Any
	Offset: typing.Any
	Size: typing.Any
	Transparency: typing.Any
		# Method 'Type' returns enumeration of type 'MsoReflectionType'
	Type: MsoReflectionType
	def __iter__(self):
		...

class Ruler2(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
		# Method 'Levels' returns object of type 'RulerLevels2'
	Levels: RulerLevels2
	Parent: typing.Any
		# Method 'TabStops' returns object of type 'TabStops2'
	TabStops: TabStops2
	def __iter__(self):
		...

class RulerLevel2(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	FirstMargin: typing.Any
	LeftMargin: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class RulerLevels2(typing.Protocol):

	# Result is of type RulerLevel2
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> RulerLevel2:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> RulerLevel2:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ScopeFolder(typing.Protocol):

	def AddToSearchFolders(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	Name: typing.Any
	Path: typing.Any
		# Method 'ScopeFolders' returns object of type 'ScopeFolders'
	ScopeFolders: ScopeFolders
	# Default property for this class is 'Name'
	def __call__(self):
		...
	def __iter__(self):
		...

class ScopeFolders(typing.Protocol):

	# Result is of type ScopeFolder
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> ScopeFolder:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> ScopeFolder:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Script(typing.Protocol):

	def Delete(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	Extended: typing.Any
	Id: typing.Any
		# Method 'Language' returns enumeration of type 'MsoScriptLanguage'
	Language: MsoScriptLanguage
		# Method 'Location' returns enumeration of type 'MsoScriptLocation'
	Location: MsoScriptLocation
	Parent: typing.Any
	ScriptText: typing.Any
	Shape: typing.Any
	# Default property for this class is 'ScriptText'
	def __call__(self):
		...
	def __iter__(self):
		...

class Scripts(typing.Protocol):

	# Result is of type Script
	def Add(self, Anchor: Dispatch=None, Location: MsoScriptLocation=2, Language: MsoScriptLanguage=2, Id: str=''
			, Extended: str='', ScriptText: str='') -> Script:
		...
	def Delete(self) -> None:
		...
	# Result is of type Script
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> Script:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> Script:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SearchFolders(typing.Protocol):

	def Add(self, ScopeFolder: ScopeFolder=defaultNamedNotOptArg) -> None:
		...
	# Result is of type ScopeFolder
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> ScopeFolder:
		...
	def Remove(self, Index: int=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> ScopeFolder:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SearchScope(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
		# Method 'ScopeFolder' returns object of type 'ScopeFolder'
	ScopeFolder: ScopeFolder
		# Method 'Type' returns enumeration of type 'MsoSearchIn'
	Type: MsoSearchIn
	# Default property for this class is 'Type'
	def __call__(self):
		...
	def __iter__(self):
		...

class SearchScopes(typing.Protocol):

	# Result is of type SearchScope
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> SearchScope:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> SearchScope:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SensitivityLabelInitInfo(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	SensitivityLabelsPolicyXml: typing.Any
	UserId: typing.Any
	# Default property for this class is 'UserId'
	def __call__(self):
		...
	def __iter__(self):
		...

class SensitivityLabelPolicy(typing.Protocol):

	def BeginInitialize(self) -> str:
		...
	def CompleteInitialize(self, SensitivityLabelInitInfo: SensitivityLabelInitInfo=defaultNamedNotOptArg) -> None:
		...
	# Result is of type SensitivityLabelInitInfo
	def CreateSensitivityLabelInitInfo(self) -> SensitivityLabelInitInfo:
		...
	Application: typing.Any
	Creator: typing.Any
		# Method 'SensitivityLabelError' returns enumeration of type 'MsoSensitivityLabelError'
	SensitivityLabelError: MsoSensitivityLabelError
	# Default method for this class is 'BeginInitialize'
	def __call__(self) -> str:
		...
	def __iter__(self):
		...

class SeriesCollection(typing.Protocol):

	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SeriesGradientStopColorFormat(typing.Protocol):

	ObjectThemeColor: typing.Any
	Parent: typing.Any
	RGB: typing.Any
	TintAndShade: typing.Any
	Transparency: typing.Any
	def __iter__(self):
		...

class SeriesGradientStopData(typing.Protocol):

	Parent: typing.Any
	StopColor: typing.Any
	StopPositionType: typing.Any
	StopValue: typing.Any
	def __iter__(self):
		...

class ServerPolicy(typing.Protocol):

	# Result is of type PolicyItem
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> PolicyItem:
		...
	Application: typing.Any
	BlockPreview: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Description: typing.Any
	Id: typing.Any
	Name: typing.Any
	Parent: typing.Any
	Statement: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> PolicyItem:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ShadowFormat(typing.Protocol):

	def IncrementOffsetX(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementOffsetY(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Blur: typing.Any
	Creator: typing.Any
		# Method 'ForeColor' returns object of type 'ColorFormat'
	ForeColor: ColorFormat
		# Method 'Obscured' returns enumeration of type 'MsoTriState'
	Obscured: MsoTriState
	OffsetX: typing.Any
	OffsetY: typing.Any
	Parent: typing.Any
		# Method 'RotateWithShape' returns enumeration of type 'MsoTriState'
	RotateWithShape: MsoTriState
	Size: typing.Any
		# Method 'Style' returns enumeration of type 'MsoShadowStyle'
	Style: MsoShadowStyle
	Transparency: typing.Any
		# Method 'Type' returns enumeration of type 'MsoShadowType'
	Type: MsoShadowType
		# Method 'Visible' returns enumeration of type 'MsoTriState'
	Visible: MsoTriState
	def __iter__(self):
		...

class Shape(typing.Protocol):

	def Apply(self) -> None:
		...
	def CanvasCropBottom(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def CanvasCropLeft(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def CanvasCropRight(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def CanvasCropTop(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def ConvertTextToSmartArt(self, Layout: SmartArtLayout=defaultNamedNotOptArg) -> None:
		...
	def Copy(self) -> None:
		...
	def Cut(self) -> None:
		...
	def Delete(self) -> None:
		...
	# Result is of type Shape
	def Duplicate(self) -> Shape:
		...
	def Flip(self, FlipCmd: MsoFlipCmd=defaultNamedNotOptArg) -> None:
		...
	def IncrementLeft(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementRotation(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementTop(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def PickUp(self) -> None:
		...
	def RerouteConnections(self) -> None:
		...
	def SaveAsPicture(self, PictureType: MsoPictureType=defaultNamedNotOptArg, FileName: str=defaultNamedNotOptArg, FSaveShapesIndividually: bool=defaultNamedNotOptArg) -> None:
		...
	def ScaleHeight(self, Factor: float=defaultNamedNotOptArg, RelativeToOriginalSize: MsoTriState=defaultNamedNotOptArg, fScale: MsoScaleFrom=0) -> None:
		...
	def ScaleWidth(self, Factor: float=defaultNamedNotOptArg, RelativeToOriginalSize: MsoTriState=defaultNamedNotOptArg, fScale: MsoScaleFrom=0) -> None:
		...
	def Select(self, Replace: typing.Any=defaultNamedOptArg) -> None:
		...
	def SetShapesDefaultProperties(self) -> None:
		...
	# Result is of type ShapeRange
	def Ungroup(self) -> ShapeRange:
		...
	def ZOrder(self, ZOrderCmd: MsoZOrderCmd=defaultNamedNotOptArg) -> None:
		...
		# Method 'Adjustments' returns object of type 'Adjustments'
	Adjustments: Adjustments
	AlternativeText: typing.Any
	Application: typing.Any
		# Method 'AutoShapeType' returns enumeration of type 'MsoAutoShapeType'
	AutoShapeType: MsoAutoShapeType
		# Method 'BackgroundStyle' returns enumeration of type 'MsoBackgroundStyleIndex'
	BackgroundStyle: MsoBackgroundStyleIndex
		# Method 'BlackWhiteMode' returns enumeration of type 'MsoBlackWhiteMode'
	BlackWhiteMode: MsoBlackWhiteMode
		# Method 'Callout' returns object of type 'CalloutFormat'
	Callout: CalloutFormat
		# Method 'CanvasItems' returns object of type 'CanvasShapes'
	CanvasItems: CanvasShapes
		# Method 'Chart' returns object of type 'IMsoChart'
	Chart: IMsoChart
		# Method 'Child' returns enumeration of type 'MsoTriState'
	Child: MsoTriState
	ConnectionSiteCount: typing.Any
		# Method 'Connector' returns enumeration of type 'MsoTriState'
	Connector: MsoTriState
		# Method 'ConnectorFormat' returns object of type 'ConnectorFormat'
	ConnectorFormat: ConnectorFormat
	Creator: typing.Any
		# Method 'Decorative' returns enumeration of type 'MsoTriState'
	Decorative: MsoTriState
		# Method 'Diagram' returns object of type 'IMsoDiagram'
	Diagram: IMsoDiagram
		# Method 'DiagramNode' returns object of type 'DiagramNode'
	DiagramNode: DiagramNode
		# Method 'Fill' returns object of type 'FillFormat'
	Fill: FillFormat
		# Method 'Glow' returns object of type 'GlowFormat'
	Glow: GlowFormat
		# Method 'GraphicStyle' returns enumeration of type 'MsoGraphicStyleIndex'
	GraphicStyle: MsoGraphicStyleIndex
		# Method 'GroupItems' returns object of type 'GroupShapes'
	GroupItems: GroupShapes
		# Method 'HasChart' returns enumeration of type 'MsoTriState'
	HasChart: MsoTriState
		# Method 'HasDiagram' returns enumeration of type 'MsoTriState'
	HasDiagram: MsoTriState
		# Method 'HasDiagramNode' returns enumeration of type 'MsoTriState'
	HasDiagramNode: MsoTriState
		# Method 'HasSmartArt' returns enumeration of type 'MsoTriState'
	HasSmartArt: MsoTriState
	Height: typing.Any
		# Method 'HorizontalFlip' returns enumeration of type 'MsoTriState'
	HorizontalFlip: MsoTriState
	Id: typing.Any
	Left: typing.Any
		# Method 'Line' returns object of type 'LineFormat'
	Line: LineFormat
		# Method 'LockAspectRatio' returns enumeration of type 'MsoTriState'
	LockAspectRatio: MsoTriState
		# Method 'Model3D' returns object of type 'Model3DFormat'
	Model3D: Model3DFormat
	Name: typing.Any
		# Method 'Nodes' returns object of type 'ShapeNodes'
	Nodes: ShapeNodes
	Parent: typing.Any
		# Method 'ParentGroup' returns object of type 'Shape'
	ParentGroup: Shape
		# Method 'PictureFormat' returns object of type 'PictureFormat'
	PictureFormat: PictureFormat
		# Method 'Reflection' returns object of type 'ReflectionFormat'
	Reflection: ReflectionFormat
	Rotation: typing.Any
		# Method 'Script' returns object of type 'Script'
	Script: Script
		# Method 'Shadow' returns object of type 'ShadowFormat'
	Shadow: ShadowFormat
		# Method 'ShapeStyle' returns enumeration of type 'MsoShapeStyleIndex'
	ShapeStyle: MsoShapeStyleIndex
		# Method 'SmartArt' returns object of type 'SmartArt'
	SmartArt: SmartArt
		# Method 'SoftEdge' returns object of type 'SoftEdgeFormat'
	SoftEdge: SoftEdgeFormat
		# Method 'TextEffect' returns object of type 'TextEffectFormat'
	TextEffect: TextEffectFormat
		# Method 'TextFrame' returns object of type 'TextFrame'
	TextFrame: TextFrame
		# Method 'TextFrame2' returns object of type 'TextFrame2'
	TextFrame2: TextFrame2
		# Method 'ThreeD' returns object of type 'ThreeDFormat'
	ThreeD: ThreeDFormat
	Title: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'MsoShapeType'
	Type: MsoShapeType
		# Method 'VerticalFlip' returns enumeration of type 'MsoTriState'
	VerticalFlip: MsoTriState
	Vertices: typing.Any
		# Method 'Visible' returns enumeration of type 'MsoTriState'
	Visible: MsoTriState
	Width: typing.Any
	ZOrderPosition: typing.Any
	def __iter__(self):
		...

class ShapeNode(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
		# Method 'EditingType' returns enumeration of type 'MsoEditingType'
	EditingType: MsoEditingType
	Parent: typing.Any
	Points: typing.Any
		# Method 'SegmentType' returns enumeration of type 'MsoSegmentType'
	SegmentType: MsoSegmentType
	def __iter__(self):
		...

class ShapeNodes(typing.Protocol):

	def Delete(self, Index: int=defaultNamedNotOptArg) -> None:
		...
	def Insert(self, Index: int=defaultNamedNotOptArg, SegmentType: MsoSegmentType=defaultNamedNotOptArg, EditingType: MsoEditingType=defaultNamedNotOptArg, X1: float=defaultNamedNotOptArg
			, Y1: float=defaultNamedNotOptArg, X2: float=0.0, Y2: float=0.0, X3: float=0.0, Y3: float=0.0) -> None:
		...
	# Result is of type ShapeNode
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> ShapeNode:
		...
	def SetEditingType(self, Index: int=defaultNamedNotOptArg, EditingType: MsoEditingType=defaultNamedNotOptArg) -> None:
		...
	def SetPosition(self, Index: int=defaultNamedNotOptArg, X1: float=defaultNamedNotOptArg, Y1: float=defaultNamedNotOptArg) -> None:
		...
	def SetSegmentType(self, Index: int=defaultNamedNotOptArg, SegmentType: MsoSegmentType=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> ShapeNode:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ShapeRange(typing.Protocol):

	def Align(self, AlignCmd: MsoAlignCmd=defaultNamedNotOptArg, RelativeTo: MsoTriState=defaultNamedNotOptArg) -> None:
		...
	def Apply(self) -> None:
		...
	def CanvasCropBottom(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def CanvasCropLeft(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def CanvasCropRight(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def CanvasCropTop(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def Copy(self) -> None:
		...
	def Cut(self) -> None:
		...
	def Delete(self) -> None:
		...
	def Distribute(self, DistributeCmd: MsoDistributeCmd=defaultNamedNotOptArg, RelativeTo: MsoTriState=defaultNamedNotOptArg) -> None:
		...
	# Result is of type ShapeRange
	def Duplicate(self) -> ShapeRange:
		...
	def Flip(self, FlipCmd: MsoFlipCmd=defaultNamedNotOptArg) -> None:
		...
	# Result is of type Shape
	def Group(self) -> Shape:
		...
	def IncrementLeft(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementRotation(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementTop(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	# Result is of type Shape
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	def MergeShapes(self, MergeCmd: MsoMergeCmd=defaultNamedNotOptArg, PrimaryShape: Shape=0) -> None:
		...
	def PickUp(self) -> None:
		...
	# Result is of type Shape
	def Regroup(self) -> Shape:
		...
	def RerouteConnections(self) -> None:
		...
	def SaveAsPicture(self, PictureType: MsoPictureType=defaultNamedNotOptArg, FileName: str=defaultNamedNotOptArg, FSaveShapesIndividually: bool=defaultNamedNotOptArg) -> None:
		...
	def ScaleHeight(self, Factor: float=defaultNamedNotOptArg, RelativeToOriginalSize: MsoTriState=defaultNamedNotOptArg, fScale: MsoScaleFrom=0) -> None:
		...
	def ScaleWidth(self, Factor: float=defaultNamedNotOptArg, RelativeToOriginalSize: MsoTriState=defaultNamedNotOptArg, fScale: MsoScaleFrom=0) -> None:
		...
	def Select(self, Replace: typing.Any=defaultNamedOptArg) -> None:
		...
	def SetShapesDefaultProperties(self) -> None:
		...
	# Result is of type ShapeRange
	def Ungroup(self) -> ShapeRange:
		...
	def ZOrder(self, ZOrderCmd: MsoZOrderCmd=defaultNamedNotOptArg) -> None:
		...
		# Method 'Adjustments' returns object of type 'Adjustments'
	Adjustments: Adjustments
	AlternativeText: typing.Any
	Application: typing.Any
		# Method 'AutoShapeType' returns enumeration of type 'MsoAutoShapeType'
	AutoShapeType: MsoAutoShapeType
		# Method 'BackgroundStyle' returns enumeration of type 'MsoBackgroundStyleIndex'
	BackgroundStyle: MsoBackgroundStyleIndex
		# Method 'BlackWhiteMode' returns enumeration of type 'MsoBlackWhiteMode'
	BlackWhiteMode: MsoBlackWhiteMode
		# Method 'Callout' returns object of type 'CalloutFormat'
	Callout: CalloutFormat
		# Method 'CanvasItems' returns object of type 'CanvasShapes'
	CanvasItems: CanvasShapes
		# Method 'Chart' returns object of type 'IMsoChart'
	Chart: IMsoChart
		# Method 'Child' returns enumeration of type 'MsoTriState'
	Child: MsoTriState
	ConnectionSiteCount: typing.Any
		# Method 'Connector' returns enumeration of type 'MsoTriState'
	Connector: MsoTriState
		# Method 'ConnectorFormat' returns object of type 'ConnectorFormat'
	ConnectorFormat: ConnectorFormat
	Count: typing.Any
	Creator: typing.Any
		# Method 'Decorative' returns enumeration of type 'MsoTriState'
	Decorative: MsoTriState
		# Method 'Diagram' returns object of type 'IMsoDiagram'
	Diagram: IMsoDiagram
		# Method 'DiagramNode' returns object of type 'DiagramNode'
	DiagramNode: DiagramNode
		# Method 'Fill' returns object of type 'FillFormat'
	Fill: FillFormat
		# Method 'Glow' returns object of type 'GlowFormat'
	Glow: GlowFormat
		# Method 'GraphicStyle' returns enumeration of type 'MsoGraphicStyleIndex'
	GraphicStyle: MsoGraphicStyleIndex
		# Method 'GroupItems' returns object of type 'GroupShapes'
	GroupItems: GroupShapes
		# Method 'HasChart' returns enumeration of type 'MsoTriState'
	HasChart: MsoTriState
		# Method 'HasDiagram' returns enumeration of type 'MsoTriState'
	HasDiagram: MsoTriState
		# Method 'HasDiagramNode' returns enumeration of type 'MsoTriState'
	HasDiagramNode: MsoTriState
	Height: typing.Any
		# Method 'HorizontalFlip' returns enumeration of type 'MsoTriState'
	HorizontalFlip: MsoTriState
	Id: typing.Any
	Left: typing.Any
		# Method 'Line' returns object of type 'LineFormat'
	Line: LineFormat
		# Method 'LockAspectRatio' returns enumeration of type 'MsoTriState'
	LockAspectRatio: MsoTriState
		# Method 'Model3D' returns object of type 'Model3DFormat'
	Model3D: Model3DFormat
	Name: typing.Any
		# Method 'Nodes' returns object of type 'ShapeNodes'
	Nodes: ShapeNodes
	Parent: typing.Any
		# Method 'ParentGroup' returns object of type 'Shape'
	ParentGroup: Shape
		# Method 'PictureFormat' returns object of type 'PictureFormat'
	PictureFormat: PictureFormat
		# Method 'Reflection' returns object of type 'ReflectionFormat'
	Reflection: ReflectionFormat
	Rotation: typing.Any
		# Method 'Script' returns object of type 'Script'
	Script: Script
		# Method 'Shadow' returns object of type 'ShadowFormat'
	Shadow: ShadowFormat
		# Method 'ShapeStyle' returns enumeration of type 'MsoShapeStyleIndex'
	ShapeStyle: MsoShapeStyleIndex
		# Method 'SoftEdge' returns object of type 'SoftEdgeFormat'
	SoftEdge: SoftEdgeFormat
		# Method 'TextEffect' returns object of type 'TextEffectFormat'
	TextEffect: TextEffectFormat
		# Method 'TextFrame' returns object of type 'TextFrame'
	TextFrame: TextFrame
		# Method 'TextFrame2' returns object of type 'TextFrame2'
	TextFrame2: TextFrame2
		# Method 'ThreeD' returns object of type 'ThreeDFormat'
	ThreeD: ThreeDFormat
	Title: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'MsoShapeType'
	Type: MsoShapeType
		# Method 'VerticalFlip' returns enumeration of type 'MsoTriState'
	VerticalFlip: MsoTriState
	Vertices: typing.Any
		# Method 'Visible' returns enumeration of type 'MsoTriState'
	Visible: MsoTriState
	Width: typing.Any
	ZOrderPosition: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Shapes(typing.Protocol):

	# Result is of type Shape
	def Add3DModel(self, FileName: str=defaultNamedNotOptArg, LinkToFile: MsoTriState=defaultNamedNotOptArg, SaveWithDocument: MsoTriState=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg
			, Top: float=defaultNamedNotOptArg, Width: float=-1.0, Height: float=-1.0) -> Shape:
		...
	# Result is of type Shape
	def AddCallout(self, Type: MsoCalloutType=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddCanvas(self, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddChart(self, Type: XlChartType=-1, Left: float=-1.0, Top: float=-1.0, Width: float=-1.0
			, Height: float=-1.0) -> Shape:
		...
	# Result is of type Shape
	def AddChart2(self, Style: int=-1, Type: XlChartType=-1, Left: float=-1.0, Top: float=-1.0
			, Width: float=-1.0, Height: float=-1.0, NewLayout: bool=True) -> Shape:
		...
	# Result is of type Shape
	def AddConnector(self, Type: MsoConnectorType=defaultNamedNotOptArg, BeginX: float=defaultNamedNotOptArg, BeginY: float=defaultNamedNotOptArg, EndX: float=defaultNamedNotOptArg
			, EndY: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddCurve(self, SafeArrayOfPoints: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddDiagram(self, Type: MsoDiagramType=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddLabel(self, Orientation: MsoTextOrientation=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddLine(self, BeginX: float=defaultNamedNotOptArg, BeginY: float=defaultNamedNotOptArg, EndX: float=defaultNamedNotOptArg, EndY: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddPicture(self, FileName: str=defaultNamedNotOptArg, LinkToFile: MsoTriState=defaultNamedNotOptArg, SaveWithDocument: MsoTriState=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg
			, Top: float=defaultNamedNotOptArg, Width: float=-1.0, Height: float=-1.0) -> Shape:
		...
	# Result is of type Shape
	def AddPicture2(self, FileName: str=defaultNamedNotOptArg, LinkToFile: MsoTriState=defaultNamedNotOptArg, SaveWithDocument: MsoTriState=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg
			, Top: float=defaultNamedNotOptArg, Width: float=-1.0, Height: float=-1.0, Compress: MsoPictureCompress=-1) -> Shape:
		...
	# Result is of type Shape
	def AddPolyline(self, SafeArrayOfPoints: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddShape(self, Type: MsoAutoShapeType=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddSmartArt(self, Layout: SmartArtLayout=defaultNamedNotOptArg, Left: float=-1.0, Top: float=-1.0, Width: float=-1.0
			, Height: float=-1.0) -> Shape:
		...
	# Result is of type Shape
	def AddTable(self, NumRows: int=defaultNamedNotOptArg, NumColumns: int=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg
			, Width: float=defaultNamedNotOptArg, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddTextEffect(self, PresetTextEffect: MsoPresetTextEffect=defaultNamedNotOptArg, Text: str=defaultNamedNotOptArg, FontName: str=defaultNamedNotOptArg, FontSize: float=defaultNamedNotOptArg
			, FontBold: MsoTriState=defaultNamedNotOptArg, FontItalic: MsoTriState=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type Shape
	def AddTextbox(self, Orientation: MsoTextOrientation=defaultNamedNotOptArg, Left: float=defaultNamedNotOptArg, Top: float=defaultNamedNotOptArg, Width: float=defaultNamedNotOptArg
			, Height: float=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type FreeformBuilder
	def BuildFreeform(self, EditingType: MsoEditingType=defaultNamedNotOptArg, X1: float=defaultNamedNotOptArg, Y1: float=defaultNamedNotOptArg) -> FreeformBuilder:
		...
	# Result is of type Shape
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	# Result is of type ShapeRange
	def Range(self, Index: typing.Any=defaultNamedNotOptArg) -> ShapeRange:
		...
	def SelectAll(self) -> None:
		...
	Application: typing.Any
		# Method 'Background' returns object of type 'Shape'
	Background: Shape
	Count: typing.Any
	Creator: typing.Any
		# Method 'Default' returns object of type 'Shape'
	Default: Shape
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> Shape:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SharedWorkspace(typing.Protocol):

	def CreateNew(self, URL: typing.Any=defaultNamedOptArg, Name: typing.Any=defaultNamedOptArg) -> None:
		...
	def Delete(self) -> None:
		...
	def Disconnect(self) -> None:
		...
	def Refresh(self) -> None:
		...
	def RemoveDocument(self) -> None:
		...
	Application: typing.Any
	Connected: typing.Any
	Creator: typing.Any
		# Method 'Files' returns object of type 'SharedWorkspaceFiles'
	Files: SharedWorkspaceFiles
		# Method 'Folders' returns object of type 'SharedWorkspaceFolders'
	Folders: SharedWorkspaceFolders
	LastRefreshed: typing.Any
		# Method 'Links' returns object of type 'SharedWorkspaceLinks'
	Links: SharedWorkspaceLinks
		# Method 'Members' returns object of type 'SharedWorkspaceMembers'
	Members: SharedWorkspaceMembers
	Name: typing.Any
	Parent: typing.Any
	SourceURL: typing.Any
		# Method 'Tasks' returns object of type 'SharedWorkspaceTasks'
	Tasks: SharedWorkspaceTasks
	URL: typing.Any
	# Default property for this class is 'Name'
	def __call__(self):
		...
	def __iter__(self):
		...

class SharedWorkspaceFile(typing.Protocol):

	def Delete(self) -> None:
		...
	Application: typing.Any
	CreatedBy: typing.Any
	CreatedDate: typing.Any
	Creator: typing.Any
	ModifiedBy: typing.Any
	ModifiedDate: typing.Any
	Parent: typing.Any
	URL: typing.Any
	# Default property for this class is 'URL'
	def __call__(self):
		...
	def __iter__(self):
		...

class SharedWorkspaceFiles(typing.Protocol):

	# Result is of type SharedWorkspaceFile
	def Add(self, FileName: str=defaultNamedNotOptArg, ParentFolder: typing.Any=defaultNamedOptArg, OverwriteIfFileAlreadyExists: typing.Any=defaultNamedOptArg, KeepInSync: typing.Any=defaultNamedOptArg) -> SharedWorkspaceFile:
		...
	# Result is of type SharedWorkspaceFile
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceFile:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	ItemCountExceeded: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceFile:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SharedWorkspaceFolder(typing.Protocol):

	def Delete(self, DeleteEventIfFolderContainsFiles: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	FolderName: typing.Any
	Parent: typing.Any
	# Default property for this class is 'FolderName'
	def __call__(self):
		...
	def __iter__(self):
		...

class SharedWorkspaceFolders(typing.Protocol):

	# Result is of type SharedWorkspaceFolder
	def Add(self, FolderName: str=defaultNamedNotOptArg, ParentFolder: typing.Any=defaultNamedOptArg) -> SharedWorkspaceFolder:
		...
	# Result is of type SharedWorkspaceFolder
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceFolder:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	ItemCountExceeded: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceFolder:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SharedWorkspaceLink(typing.Protocol):

	def Delete(self) -> None:
		...
	def Save(self) -> None:
		...
	Application: typing.Any
	CreatedBy: typing.Any
	CreatedDate: typing.Any
	Creator: typing.Any
	Description: typing.Any
	ModifiedBy: typing.Any
	ModifiedDate: typing.Any
	Notes: typing.Any
	Parent: typing.Any
	URL: typing.Any
	# Default property for this class is 'URL'
	def __call__(self):
		...
	def __iter__(self):
		...

class SharedWorkspaceLinks(typing.Protocol):

	# Result is of type SharedWorkspaceLink
	def Add(self, URL: str=defaultNamedNotOptArg, Description: typing.Any=defaultNamedOptArg, Notes: typing.Any=defaultNamedOptArg) -> SharedWorkspaceLink:
		...
	# Result is of type SharedWorkspaceLink
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceLink:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	ItemCountExceeded: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceLink:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SharedWorkspaceMember(typing.Protocol):

	def Delete(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	DomainName: typing.Any
	Email: typing.Any
	Id: typing.Any
	Name: typing.Any
	Parent: typing.Any
	# Default property for this class is 'DomainName'
	def __call__(self):
		...
	def __iter__(self):
		...

class SharedWorkspaceMembers(typing.Protocol):

	# Result is of type SharedWorkspaceMember
	def Add(self, Email: str=defaultNamedNotOptArg, DomainName: str=defaultNamedNotOptArg, DisplayName: str=defaultNamedNotOptArg, Role: typing.Any=defaultNamedOptArg) -> SharedWorkspaceMember:
		...
	# Result is of type SharedWorkspaceMember
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceMember:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	ItemCountExceeded: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceMember:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SharedWorkspaceTask(typing.Protocol):

	def Delete(self) -> None:
		...
	def Save(self) -> None:
		...
	Application: typing.Any
	AssignedTo: typing.Any
	CreatedBy: typing.Any
	CreatedDate: typing.Any
	Creator: typing.Any
	Description: typing.Any
	DueDate: typing.Any
	ModifiedBy: typing.Any
	ModifiedDate: typing.Any
	Parent: typing.Any
		# Method 'Priority' returns enumeration of type 'MsoSharedWorkspaceTaskPriority'
	Priority: MsoSharedWorkspaceTaskPriority
		# Method 'Status' returns enumeration of type 'MsoSharedWorkspaceTaskStatus'
	Status: MsoSharedWorkspaceTaskStatus
	Title: typing.Any
	# Default property for this class is 'Title'
	def __call__(self):
		...
	def __iter__(self):
		...

class SharedWorkspaceTasks(typing.Protocol):

	# Result is of type SharedWorkspaceTask
	def Add(self, Title: str=defaultNamedNotOptArg, Status: typing.Any=defaultNamedOptArg, Priority: typing.Any=defaultNamedOptArg, Assignee: typing.Any=defaultNamedOptArg
			, Description: typing.Any=defaultNamedOptArg, DueDate: typing.Any=defaultNamedOptArg) -> SharedWorkspaceTask:
		...
	# Result is of type SharedWorkspaceTask
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceTask:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	ItemCountExceeded: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> SharedWorkspaceTask:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Signature(typing.Protocol):

	def Delete(self) -> None:
		...
	def ShowDetails(self) -> None:
		...
	def Sign(self, varSigImg: typing.Any=defaultNamedOptArg, varDelSuggSigner: typing.Any=defaultNamedOptArg, varDelSuggSignerLine2: typing.Any=defaultNamedOptArg, varDelSuggSignerEmail: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
	AttachCertificate: typing.Any
	CanSetup: typing.Any
	Creator: typing.Any
		# Method 'Details' returns object of type 'SignatureInfo'
	Details: SignatureInfo
	ExpireDate: typing.Any
	IsCertificateExpired: typing.Any
	IsCertificateRevoked: typing.Any
	IsSignatureLine: typing.Any
	IsSigned: typing.Any
	IsValid: typing.Any
	Issuer: typing.Any
	Parent: typing.Any
		# Method 'Setup' returns object of type 'SignatureSetup'
	Setup: SignatureSetup
	SignDate: typing.Any
	SignatureLineShape: typing.Any
	Signer: typing.Any
	SortHint: typing.Any
	def __iter__(self):
		...

class SignatureInfo(typing.Protocol):

	def GetCertificateDetail(self, certdet: CertificateDetail=defaultNamedNotOptArg) -> typing.Any:
		...
	def GetSignatureDetail(self, sigdet: SignatureDetail=defaultNamedNotOptArg) -> typing.Any:
		...
	def SelectCertificateDetailByThumbprint(self, bstrThumbprint: str=defaultNamedNotOptArg) -> None:
		...
	def SelectSignatureCertificate(self, ParentWindow: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def ShowSignatureCertificate(self, ParentWindow: typing.Any=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
		# Method 'CertificateVerificationResults' returns enumeration of type 'CertificateVerificationResults'
	CertificateVerificationResults: CertificateVerificationResults
		# Method 'ContentVerificationResults' returns enumeration of type 'ContentVerificationResults'
	ContentVerificationResults: ContentVerificationResults
	Creator: typing.Any
	IsCertificateExpired: typing.Any
	IsCertificateRevoked: typing.Any
	IsCertificateUntrusted: typing.Any
	IsValid: typing.Any
	ReadOnly: typing.Any
	SignatureComment: typing.Any
		# Method 'SignatureImage' returns object of type 'Picture'
	SignatureImage: Picture
	SignatureProvider: typing.Any
	SignatureText: typing.Any
	def __iter__(self):
		...

class SignatureProvider(typing.Protocol):

	# Result is of type Picture
	def GenerateSignatureLineImage(self, siglnimg: SignatureLineImage=defaultNamedNotOptArg, psigsetup: SignatureSetup=defaultNamedNotOptArg, psiginfo: SignatureInfo=defaultNamedNotOptArg, XmlDsigStream: typing.Any=defaultNamedNotOptArg) -> Picture:
		...
	def GetProviderDetail(self, sigprovdet: SignatureProviderDetail=defaultNamedNotOptArg) -> typing.Any:
		...
	def HashStream(self, QueryContinue: typing.Any=defaultNamedNotOptArg, Stream: typing.Any=defaultNamedNotOptArg) -> typing.List[int]:
		...
	def NotifySignatureAdded(self, ParentWindow: typing.Any=defaultNamedNotOptArg, psigsetup: SignatureSetup=defaultNamedNotOptArg, psiginfo: SignatureInfo=defaultNamedNotOptArg) -> None:
		...
	def ShowSignatureDetails(self, ParentWindow: typing.Any=defaultNamedNotOptArg, psigsetup: SignatureSetup=defaultNamedNotOptArg, psiginfo: SignatureInfo=defaultNamedNotOptArg, XmlDsigStream: typing.Any=defaultNamedNotOptArg
			, pcontverres: ContentVerificationResults=defaultNamedNotOptArg, pcertverres: CertificateVerificationResults=defaultNamedNotOptArg) -> None:
		...
	def ShowSignatureSetup(self, ParentWindow: typing.Any=defaultNamedNotOptArg, psigsetup: SignatureSetup=defaultNamedNotOptArg) -> None:
		...
	def ShowSigningCeremony(self, ParentWindow: typing.Any=defaultNamedNotOptArg, psigsetup: SignatureSetup=defaultNamedNotOptArg, psiginfo: SignatureInfo=defaultNamedNotOptArg) -> None:
		...
	def SignXmlDsig(self, QueryContinue: typing.Any=defaultNamedNotOptArg, psigsetup: SignatureSetup=defaultNamedNotOptArg, psiginfo: SignatureInfo=defaultNamedNotOptArg, XmlDsigStream: typing.Any=defaultNamedNotOptArg) -> None:
		...
	def VerifyXmlDsig(self, QueryContinue: typing.Any=defaultNamedNotOptArg, psigsetup: SignatureSetup=defaultNamedNotOptArg, psiginfo: SignatureInfo=defaultNamedNotOptArg, XmlDsigStream: typing.Any=defaultNamedNotOptArg
			, pcontverres: ContentVerificationResults=defaultNamedNotOptArg, pcertverres: CertificateVerificationResults=defaultNamedNotOptArg) -> None:
		...
	def __iter__(self):
		...

class SignatureSet(typing.Protocol):

	# Result is of type Signature
	def Add(self) -> Signature:
		...
	# Result is of type Signature
	def AddNonVisibleSignature(self, varSigProv: typing.Any=defaultNamedOptArg) -> Signature:
		...
	# Result is of type Signature
	def AddSignatureLine(self, varSigProv: typing.Any=defaultNamedOptArg) -> Signature:
		...
	def Commit(self) -> None:
		...
	# Result is of type Signature
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, iSig: int=defaultNamedNotOptArg) -> Signature:
		...
	Application: typing.Any
	CanAddSignatureLine: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
		# Method 'Subset' returns enumeration of type 'MsoSignatureSubset'
	Subset: MsoSignatureSubset
	# Default method for this class is 'Item'
	def __call__(self, iSig: int=defaultNamedNotOptArg) -> Signature:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SignatureSetup(typing.Protocol):

	AdditionalXml: typing.Any
	AllowComments: typing.Any
	Application: typing.Any
	Creator: typing.Any
	Id: typing.Any
	ReadOnly: typing.Any
	ShowSignDate: typing.Any
	SignatureProvider: typing.Any
	SigningInstructions: typing.Any
	SuggestedSigner: typing.Any
	SuggestedSignerEmail: typing.Any
	SuggestedSignerLine2: typing.Any
	def __iter__(self):
		...

class SmartArt(typing.Protocol):

	def Reset(self) -> None:
		...
		# Method 'AllNodes' returns object of type 'SmartArtNodes'
	AllNodes: SmartArtNodes
	Application: typing.Any
		# Method 'Color' returns object of type 'SmartArtColor'
	Color: SmartArtColor
	Creator: typing.Any
		# Method 'Layout' returns object of type 'SmartArtLayout'
	Layout: SmartArtLayout
		# Method 'Nodes' returns object of type 'SmartArtNodes'
	Nodes: SmartArtNodes
	Parent: typing.Any
		# Method 'QuickStyle' returns object of type 'SmartArtQuickStyle'
	QuickStyle: SmartArtQuickStyle
		# Method 'Reverse' returns enumeration of type 'MsoTriState'
	Reverse: MsoTriState
	def __iter__(self):
		...

class SmartArtColor(typing.Protocol):

	Application: typing.Any
	Category: typing.Any
	Creator: typing.Any
	Description: typing.Any
	Id: typing.Any
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class SmartArtColors(typing.Protocol):

	# Result is of type SmartArtColor
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> SmartArtColor:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> SmartArtColor:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SmartArtLayout(typing.Protocol):

	Application: typing.Any
	Category: typing.Any
	Creator: typing.Any
	Description: typing.Any
	Id: typing.Any
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class SmartArtLayouts(typing.Protocol):

	# Result is of type SmartArtLayout
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> SmartArtLayout:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> SmartArtLayout:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SmartArtNode(typing.Protocol):

	# Result is of type SmartArtNode
	def AddNode(self, Position: MsoSmartArtNodePosition=1, Type: MsoSmartArtNodeType=1) -> SmartArtNode:
		...
	def Delete(self) -> None:
		...
	def Demote(self) -> None:
		...
	def Larger(self) -> None:
		...
	def Promote(self) -> None:
		...
	def ReorderDown(self) -> None:
		...
	def ReorderUp(self) -> None:
		...
	def Smaller(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
		# Method 'Hidden' returns enumeration of type 'MsoTriState'
	Hidden: MsoTriState
	Level: typing.Any
		# Method 'Nodes' returns object of type 'SmartArtNodes'
	Nodes: SmartArtNodes
		# Method 'OrgChartLayout' returns enumeration of type 'MsoOrgChartLayoutType'
	OrgChartLayout: MsoOrgChartLayoutType
	Parent: typing.Any
		# Method 'ParentNode' returns object of type 'SmartArtNode'
	ParentNode: SmartArtNode
		# Method 'Shapes' returns object of type 'ShapeRange'
	Shapes: ShapeRange
		# Method 'TextFrame2' returns object of type 'TextFrame2'
	TextFrame2: TextFrame2
		# Method 'Type' returns enumeration of type 'MsoSmartArtNodeType'
	Type: MsoSmartArtNodeType
	def __iter__(self):
		...

class SmartArtNodes(typing.Protocol):

	# Result is of type SmartArtNode
	def Add(self) -> SmartArtNode:
		...
	# Result is of type SmartArtNode
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> SmartArtNode:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> SmartArtNode:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SmartArtQuickStyle(typing.Protocol):

	Application: typing.Any
	Category: typing.Any
	Creator: typing.Any
	Description: typing.Any
	Id: typing.Any
	Name: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class SmartArtQuickStyles(typing.Protocol):

	# Result is of type SmartArtQuickStyle
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> SmartArtQuickStyle:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> SmartArtQuickStyle:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class SmartDocument(typing.Protocol):

	def PickSolution(self, ConsiderAllSchemas: bool=False) -> None:
		...
	def RefreshPane(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	SolutionID: typing.Any
	SolutionURL: typing.Any
	def __iter__(self):
		...

class SoftEdgeFormat(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Radius: typing.Any
		# Method 'Type' returns enumeration of type 'MsoSoftEdgeType'
	Type: MsoSoftEdgeType
	def __iter__(self):
		...

class Sync(typing.Protocol):

	def GetUpdate(self) -> None:
		...
	def OpenVersion(self, SyncVersionType: MsoSyncVersionType=defaultNamedNotOptArg) -> None:
		...
	def PutUpdate(self) -> None:
		...
	def ResolveConflict(self, SyncConflictResolution: MsoSyncConflictResolutionType=defaultNamedNotOptArg) -> None:
		...
	def Unsuspend(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
		# Method 'ErrorType' returns enumeration of type 'MsoSyncErrorType'
	ErrorType: MsoSyncErrorType
	LastSyncTime: typing.Any
	Parent: typing.Any
		# Method 'Status' returns enumeration of type 'MsoSyncStatusType'
	Status: MsoSyncStatusType
	WorkspaceLastChangedBy: typing.Any
	# Default property for this class is 'Status'
	def __call__(self):
		...
	def __iter__(self):
		...

class TabStop2(typing.Protocol):

	def Clear(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	Position: typing.Any
		# Method 'Type' returns enumeration of type 'MsoTabStopType'
	Type: MsoTabStopType
	def __iter__(self):
		...

class TabStops2(typing.Protocol):

	# Result is of type TabStop2
	def Add(self, Type: MsoTabStopType=defaultNamedNotOptArg, Position: float=defaultNamedNotOptArg) -> TabStop2:
		...
	# Result is of type TabStop2
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> TabStop2:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	DefaultSpacing: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> TabStop2:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class TextColumn2(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Number: typing.Any
	Spacing: typing.Any
		# Method 'TextDirection' returns enumeration of type 'MsoTextDirection'
	TextDirection: MsoTextDirection
	def __iter__(self):
		...

class TextEffectFormat(typing.Protocol):

	def ToggleVerticalText(self) -> None:
		...
		# Method 'Alignment' returns enumeration of type 'MsoTextEffectAlignment'
	Alignment: MsoTextEffectAlignment
	Application: typing.Any
	Creator: typing.Any
		# Method 'FontBold' returns enumeration of type 'MsoTriState'
	FontBold: MsoTriState
		# Method 'FontItalic' returns enumeration of type 'MsoTriState'
	FontItalic: MsoTriState
	FontName: typing.Any
	FontSize: typing.Any
		# Method 'KernedPairs' returns enumeration of type 'MsoTriState'
	KernedPairs: MsoTriState
		# Method 'NormalizedHeight' returns enumeration of type 'MsoTriState'
	NormalizedHeight: MsoTriState
	Parent: typing.Any
		# Method 'PresetShape' returns enumeration of type 'MsoPresetTextEffectShape'
	PresetShape: MsoPresetTextEffectShape
		# Method 'PresetTextEffect' returns enumeration of type 'MsoPresetTextEffect'
	PresetTextEffect: MsoPresetTextEffect
		# Method 'RotatedChars' returns enumeration of type 'MsoTriState'
	RotatedChars: MsoTriState
	Text: typing.Any
	Tracking: typing.Any
	def __iter__(self):
		...

class TextFrame(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	MarginBottom: typing.Any
	MarginLeft: typing.Any
	MarginRight: typing.Any
	MarginTop: typing.Any
		# Method 'Orientation' returns enumeration of type 'MsoTextOrientation'
	Orientation: MsoTextOrientation
	Parent: typing.Any
	def __iter__(self):
		...

class TextFrame2(typing.Protocol):

	def DeleteText(self) -> None:
		...
	Application: typing.Any
		# Method 'AutoSize' returns enumeration of type 'MsoAutoSize'
	AutoSize: MsoAutoSize
		# Method 'Column' returns object of type 'TextColumn2'
	Column: TextColumn2
	Creator: typing.Any
		# Method 'HasText' returns enumeration of type 'MsoTriState'
	HasText: MsoTriState
		# Method 'HorizontalAnchor' returns enumeration of type 'MsoHorizontalAnchor'
	HorizontalAnchor: MsoHorizontalAnchor
	MarginBottom: typing.Any
	MarginLeft: typing.Any
	MarginRight: typing.Any
	MarginTop: typing.Any
		# Method 'NoTextRotation' returns enumeration of type 'MsoTriState'
	NoTextRotation: MsoTriState
		# Method 'Orientation' returns enumeration of type 'MsoTextOrientation'
	Orientation: MsoTextOrientation
	Parent: typing.Any
		# Method 'PathFormat' returns enumeration of type 'MsoPathFormat'
	PathFormat: MsoPathFormat
		# Method 'Ruler' returns object of type 'Ruler2'
	Ruler: Ruler2
		# Method 'TextRange' returns object of type 'TextRange2'
	TextRange: TextRange2
		# Method 'ThreeD' returns object of type 'ThreeDFormat'
	ThreeD: ThreeDFormat
		# Method 'VerticalAnchor' returns enumeration of type 'MsoVerticalAnchor'
	VerticalAnchor: MsoVerticalAnchor
		# Method 'WarpFormat' returns enumeration of type 'MsoWarpFormat'
	WarpFormat: MsoWarpFormat
		# Method 'WordArtformat' returns enumeration of type 'MsoPresetTextEffect'
	WordArtformat: MsoPresetTextEffect
		# Method 'WordWrap' returns enumeration of type 'MsoTriState'
	WordWrap: MsoTriState
	def __iter__(self):
		...

class TextRange2(typing.Protocol):

	def AddPeriods(self) -> None:
		...
	def ChangeCase(self, Type: MsoTextChangeCase=defaultNamedNotOptArg) -> None:
		...
	def Copy(self) -> None:
		...
	def Cut(self) -> None:
		...
	def Delete(self) -> None:
		...
	# Result is of type TextRange2
	def Find(self, FindWhat: str=defaultNamedNotOptArg, After: int=0, MatchCase: MsoTriState=0, WholeWords: MsoTriState=0) -> TextRange2:
		...
	# Result is of type TextRange2
	# The method GetCharacters is actually a property, but must be used as a method to correctly pass the arguments
	def GetCharacters(self, Start: int=-1, Length: int=-1) -> TextRange2:
		...
	# Result is of type TextRange2
	# The method GetLines is actually a property, but must be used as a method to correctly pass the arguments
	def GetLines(self, Start: int=-1, Length: int=-1) -> TextRange2:
		...
	# Result is of type TextRange2
	# The method GetMathZones is actually a property, but must be used as a method to correctly pass the arguments
	def GetMathZones(self, Start: int=-1, Length: int=-1) -> TextRange2:
		...
	# Result is of type TextRange2
	# The method GetParagraphs is actually a property, but must be used as a method to correctly pass the arguments
	def GetParagraphs(self, Start: int=-1, Length: int=-1) -> TextRange2:
		...
	# Result is of type TextRange2
	# The method GetRuns is actually a property, but must be used as a method to correctly pass the arguments
	def GetRuns(self, Start: int=-1, Length: int=-1) -> TextRange2:
		...
	# Result is of type TextRange2
	# The method GetSentences is actually a property, but must be used as a method to correctly pass the arguments
	def GetSentences(self, Start: int=-1, Length: int=-1) -> TextRange2:
		...
	# Result is of type TextRange2
	# The method GetWords is actually a property, but must be used as a method to correctly pass the arguments
	def GetWords(self, Start: int=-1, Length: int=-1) -> TextRange2:
		...
	# Result is of type TextRange2
	def InsertAfter(self, NewText: str='') -> TextRange2:
		...
	# Result is of type TextRange2
	def InsertBefore(self, NewText: str='') -> TextRange2:
		...
	# Result is of type TextRange2
	def InsertChartField(self, ChartFieldType: MsoChartFieldType=defaultNamedNotOptArg, Formula: str='', Position: int=-1) -> TextRange2:
		...
	# Result is of type TextRange2
	def InsertSymbol(self, FontName: str=defaultNamedNotOptArg, CharNumber: int=defaultNamedNotOptArg, Unicode: MsoTriState=0) -> TextRange2:
		...
	# Result is of type TextRange2
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> TextRange2:
		...
	def LtrRun(self) -> None:
		...
	# Result is of type TextRange2
	def Paste(self) -> TextRange2:
		...
	# Result is of type TextRange2
	def PasteSpecial(self, Format: MsoClipboardFormat=defaultNamedNotOptArg) -> TextRange2:
		...
	def RemovePeriods(self) -> None:
		...
	# Result is of type TextRange2
	def Replace(self, FindWhat: str=defaultNamedNotOptArg, ReplaceWhat: str=defaultNamedNotOptArg, After: int=0, MatchCase: MsoTriState=0
			, WholeWords: MsoTriState=0) -> TextRange2:
		...
	def RotatedBounds(self, X1: float=pythoncom.Missing, Y1: float=pythoncom.Missing, X2: float=pythoncom.Missing, Y2: float=pythoncom.Missing
			, X3: float=pythoncom.Missing, Y3: float=pythoncom.Missing, x4: float=pythoncom.Missing, y4: float=pythoncom.Missing) -> None:
		...
	def RtlRun(self) -> None:
		...
	def Select(self) -> None:
		...
	# Result is of type TextRange2
	def TrimText(self) -> TextRange2:
		...
	Application: typing.Any
	BoundHeight: typing.Any
	BoundLeft: typing.Any
	BoundTop: typing.Any
	BoundWidth: typing.Any
		# Method 'Characters' returns object of type 'TextRange2'
	Characters: TextRange2
	Count: typing.Any
	Creator: typing.Any
		# Method 'Font' returns object of type 'Font2'
	Font: Font2
		# Method 'LanguageID' returns enumeration of type 'MsoLanguageID'
	LanguageID: MsoLanguageID
	Length: typing.Any
		# Method 'Lines' returns object of type 'TextRange2'
	Lines: TextRange2
		# Method 'MathZones' returns object of type 'TextRange2'
	MathZones: TextRange2
		# Method 'ParagraphFormat' returns object of type 'ParagraphFormat2'
	ParagraphFormat: ParagraphFormat2
		# Method 'Paragraphs' returns object of type 'TextRange2'
	Paragraphs: TextRange2
	Parent: typing.Any
		# Method 'Runs' returns object of type 'TextRange2'
	Runs: TextRange2
		# Method 'Sentences' returns object of type 'TextRange2'
	Sentences: TextRange2
	Start: typing.Any
	Text: typing.Any
		# Method 'Words' returns object of type 'TextRange2'
	Words: TextRange2
	# Default property for this class is 'Text'
	def __call__(self):
		...
	def __iter__(self):
		...
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ThemeColor(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	RGB: typing.Any
		# Method 'ThemeColorSchemeIndex' returns enumeration of type 'MsoThemeColorSchemeIndex'
	ThemeColorSchemeIndex: MsoThemeColorSchemeIndex
	# Default property for this class is 'RGB'
	def __call__(self):
		...
	def __iter__(self):
		...

class ThemeColorScheme(typing.Protocol):

	# Result is of type ThemeColor
	def Colors(self, Index: MsoThemeColorSchemeIndex=defaultNamedNotOptArg) -> ThemeColor:
		...
	def GetCustomColor(self, Name: str=defaultNamedNotOptArg) -> int:
		...
	def Load(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	def Save(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Colors'
	def __call__(self, Index: MsoThemeColorSchemeIndex=defaultNamedNotOptArg) -> ThemeColor:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ThemeEffectScheme(typing.Protocol):

	def Load(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	def __iter__(self):
		...

class ThemeFont(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	Name: typing.Any
	Parent: typing.Any
	# Default property for this class is 'Name'
	def __call__(self):
		...
	def __iter__(self):
		...

class ThemeFontScheme(typing.Protocol):

	def Load(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	def Save(self, FileName: str=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
		# Method 'MajorFont' returns object of type 'ThemeFonts'
	MajorFont: ThemeFonts
		# Method 'MinorFont' returns object of type 'ThemeFonts'
	MinorFont: ThemeFonts
	Parent: typing.Any
	def __iter__(self):
		...

class ThemeFonts(typing.Protocol):

	# Result is of type ThemeFont
	def Item(self, Index: MsoFontLanguageIndex=defaultNamedNotOptArg) -> ThemeFont:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: MsoFontLanguageIndex=defaultNamedNotOptArg) -> ThemeFont:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class ThreeDFormat(typing.Protocol):

	def IncrementRotationHorizontal(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementRotationVertical(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementRotationX(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementRotationY(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def IncrementRotationZ(self, Increment: float=defaultNamedNotOptArg) -> None:
		...
	def ResetRotation(self) -> None:
		...
	def SetExtrusionDirection(self, PresetExtrusionDirection: MsoPresetExtrusionDirection=defaultNamedNotOptArg) -> None:
		...
	def SetPresetCamera(self, PresetCamera: MsoPresetCamera=defaultNamedNotOptArg) -> None:
		...
	def SetThreeDFormat(self, PresetThreeDFormat: MsoPresetThreeDFormat=defaultNamedNotOptArg) -> None:
		...
	Application: typing.Any
	BevelBottomDepth: typing.Any
	BevelBottomInset: typing.Any
		# Method 'BevelBottomType' returns enumeration of type 'MsoBevelType'
	BevelBottomType: MsoBevelType
	BevelTopDepth: typing.Any
	BevelTopInset: typing.Any
		# Method 'BevelTopType' returns enumeration of type 'MsoBevelType'
	BevelTopType: MsoBevelType
		# Method 'ContourColor' returns object of type 'ColorFormat'
	ContourColor: ColorFormat
	ContourWidth: typing.Any
	Creator: typing.Any
	Depth: typing.Any
		# Method 'ExtrusionColor' returns object of type 'ColorFormat'
	ExtrusionColor: ColorFormat
		# Method 'ExtrusionColorType' returns enumeration of type 'MsoExtrusionColorType'
	ExtrusionColorType: MsoExtrusionColorType
	FieldOfView: typing.Any
	LightAngle: typing.Any
	Parent: typing.Any
		# Method 'Perspective' returns enumeration of type 'MsoTriState'
	Perspective: MsoTriState
		# Method 'PresetCamera' returns enumeration of type 'MsoPresetCamera'
	PresetCamera: MsoPresetCamera
		# Method 'PresetExtrusionDirection' returns enumeration of type 'MsoPresetExtrusionDirection'
	PresetExtrusionDirection: MsoPresetExtrusionDirection
		# Method 'PresetLighting' returns enumeration of type 'MsoLightRigType'
	PresetLighting: MsoLightRigType
		# Method 'PresetLightingDirection' returns enumeration of type 'MsoPresetLightingDirection'
	PresetLightingDirection: MsoPresetLightingDirection
		# Method 'PresetLightingSoftness' returns enumeration of type 'MsoPresetLightingSoftness'
	PresetLightingSoftness: MsoPresetLightingSoftness
		# Method 'PresetMaterial' returns enumeration of type 'MsoPresetMaterial'
	PresetMaterial: MsoPresetMaterial
		# Method 'PresetThreeDFormat' returns enumeration of type 'MsoPresetThreeDFormat'
	PresetThreeDFormat: MsoPresetThreeDFormat
		# Method 'ProjectText' returns enumeration of type 'MsoTriState'
	ProjectText: MsoTriState
	RotationX: typing.Any
	RotationY: typing.Any
	RotationZ: typing.Any
		# Method 'Visible' returns enumeration of type 'MsoTriState'
	Visible: MsoTriState
	Z: typing.Any
	def __iter__(self):
		...

class Trendlines(typing.Protocol):

	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	_Default: typing.Any
	# Default property for this class is '_Default'
	def __call__(self):
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class UserPermission(typing.Protocol):

	def Remove(self) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	ExpirationDate: typing.Any
	Parent: typing.Any
	Permission: typing.Any
	UserId: typing.Any
	# Default property for this class is 'UserId'
	def __call__(self):
		...
	def __iter__(self):
		...

class WebComponent(typing.Protocol):

	def Commit(self) -> None:
		...
	def Revert(self) -> None:
		...
	def SetPlaceHolderGraphic(self, PlaceHolderGraphic: str=defaultNamedNotOptArg) -> None:
		...
	HTML: typing.Any
	Height: typing.Any
	Name: typing.Any
	Shape: typing.Any
	URL: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class WebComponentFormat(typing.Protocol):

	def LaunchPropertiesWindow(self) -> None:
		...
	Application: typing.Any
	HTML: typing.Any
	Height: typing.Any
	Name: typing.Any
	Parent: typing.Any
	PreviewGraphic: typing.Any
	URL: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class WebComponentProperties(typing.Protocol):

	HTML: typing.Any
	Height: typing.Any
	Name: typing.Any
	PreviewGraphic: typing.Any
	PreviewHTML: typing.Any
	Shape: typing.Any
	Tag: typing.Any
	URL: typing.Any
	Width: typing.Any
	def __iter__(self):
		...

class WebComponentWindowExternal(typing.Protocol):

	def CloseWindow(self) -> None:
		...
	Application: typing.Any
	ApplicationName: typing.Any
	ApplicationVersion: typing.Any
	InterfaceVersion: typing.Any
		# Method 'WebComponent' returns object of type 'WebComponent'
	WebComponent: WebComponent
	def __iter__(self):
		...

class WebPageFont(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	FixedWidthFont: typing.Any
	FixedWidthFontSize: typing.Any
	ProportionalFont: typing.Any
	ProportionalFontSize: typing.Any
	def __iter__(self):
		...

class WebPageFonts(typing.Protocol):

	# Result is of type WebPageFont
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: MsoCharacterSet=defaultNamedNotOptArg) -> WebPageFont:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: MsoCharacterSet=defaultNamedNotOptArg) -> WebPageFont:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class WorkflowTask(typing.Protocol):

	def Show(self) -> int:
		...
	Application: typing.Any
	AssignedTo: typing.Any
	CreatedBy: typing.Any
	CreatedDate: typing.Any
	Creator: typing.Any
	Description: typing.Any
	DueDate: typing.Any
	Id: typing.Any
	ListID: typing.Any
	Name: typing.Any
	WorkflowID: typing.Any
	def __iter__(self):
		...

class WorkflowTasks(typing.Protocol):

	# Result is of type WorkflowTask
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> WorkflowTask:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> WorkflowTask:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class WorkflowTemplate(typing.Protocol):

	def Show(self) -> int:
		...
	Application: typing.Any
	Creator: typing.Any
	Description: typing.Any
	DocumentLibraryName: typing.Any
	DocumentLibraryURL: typing.Any
	Id: typing.Any
	Name: typing.Any
	def __iter__(self):
		...

class WorkflowTemplates(typing.Protocol):

	# Result is of type WorkflowTemplate
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: int=defaultNamedNotOptArg) -> WorkflowTemplate:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: int=defaultNamedNotOptArg) -> WorkflowTemplate:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class _CommandBarActiveX(typing.Protocol):

	# Result is of type CommandBarControl
	def Copy(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def Delete(self, Temporary: typing.Any=defaultNamedOptArg) -> None:
		...
	def EnsureControl(self) -> None:
		...
	def Execute(self) -> None:
		...
	# The method GetaccDefaultAction is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccDescription is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDescription(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelp is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelp(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelpTopic is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelpTopic(self, pszHelpFile: str=pythoncom.Missing, varChild: typing.Any=defaultNamedOptArg) -> int:
		...
	# The method GetaccKeyboardShortcut is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccKeyboardShortcut(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccName(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccRole is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccRole(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccState is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccState(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccValue(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# Result is of type CommandBarControl
	def Move(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	# The method QueryControlInterface is actually a property, but must be used as a method to correctly pass the arguments
	def QueryControlInterface(self, bstrIid: str=defaultNamedNotOptArg) -> typing.Any:
		...
	def Reserved1(self) -> None:
		...
	def Reserved2(self) -> None:
		...
	def Reserved3(self) -> None:
		...
	def Reserved4(self) -> None:
		...
	def Reserved5(self) -> None:
		...
	def Reserved6(self) -> None:
		...
	def Reserved7(self) -> None:
		...
	def Reset(self) -> None:
		...
	def SetFocus(self) -> None:
		...
	def SetInnerObjectFactory(self, pUnk: typing.Any=defaultNamedNotOptArg) -> None:
		...
	# The method SetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccName(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccValue(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method accChild is actually a property, but must be used as a method to correctly pass the arguments
	def accChild(self, varChild: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	def accDoDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accHitTest(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> typing.Any:
		...
	def accLocation(self, pxLeft: int=pythoncom.Missing, pyTop: int=pythoncom.Missing, pcxWidth: int=pythoncom.Missing, pcyHeight: int=pythoncom.Missing
			, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accNavigate(self, navDir: int=defaultNamedNotOptArg, varStart: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def accSelect(self, flagsSelect: int=defaultNamedNotOptArg, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
	BeginGroup: typing.Any
	BuiltIn: typing.Any
	Caption: typing.Any
	Control: typing.Any
	ControlCLSID: typing.Any
	Creator: typing.Any
	DescriptionText: typing.Any
	Enabled: typing.Any
	Height: typing.Any
	HelpContextId: typing.Any
	HelpFile: typing.Any
	Id: typing.Any
	Index: typing.Any
	InstanceId: typing.Any
	InstanceIdPtr: typing.Any
	IsPriorityDropped: typing.Any
	Left: typing.Any
		# Method 'OLEUsage' returns enumeration of type 'MsoControlOLEUsage'
	OLEUsage: MsoControlOLEUsage
	OnAction: typing.Any
	Parameter: typing.Any
		# Method 'Parent' returns object of type 'CommandBar'
	Parent: CommandBar
	Priority: typing.Any
	Tag: typing.Any
	TooltipText: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'MsoControlType'
	Type: MsoControlType
	Visible: typing.Any
	Width: typing.Any
	accChildCount: typing.Any
	accDefaultAction: typing.Any
	accDescription: typing.Any
	accFocus: typing.Any
	accHelp: typing.Any
	accHelpTopic: typing.Any
	accKeyboardShortcut: typing.Any
	accName: typing.Any
	accParent: typing.Any
	accRole: typing.Any
	accSelection: typing.Any
	accState: typing.Any
	accValue: typing.Any
	def __iter__(self):
		...

class _CommandBarButton(typing.Protocol):

	# Result is of type CommandBarControl
	def Copy(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def CopyFace(self) -> None:
		...
	def Delete(self, Temporary: typing.Any=defaultNamedOptArg) -> None:
		...
	def Execute(self) -> None:
		...
	# The method GetaccDefaultAction is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccDescription is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDescription(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelp is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelp(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelpTopic is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelpTopic(self, pszHelpFile: str=pythoncom.Missing, varChild: typing.Any=defaultNamedOptArg) -> int:
		...
	# The method GetaccKeyboardShortcut is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccKeyboardShortcut(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccName(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccRole is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccRole(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccState is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccState(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccValue(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# Result is of type CommandBarControl
	def Move(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def PasteFace(self) -> None:
		...
	def Reserved1(self) -> None:
		...
	def Reserved2(self) -> None:
		...
	def Reserved3(self) -> None:
		...
	def Reserved4(self) -> None:
		...
	def Reserved5(self) -> None:
		...
	def Reserved6(self) -> None:
		...
	def Reserved7(self) -> None:
		...
	def Reset(self) -> None:
		...
	def SetFocus(self) -> None:
		...
	# The method SetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccName(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccValue(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method accChild is actually a property, but must be used as a method to correctly pass the arguments
	def accChild(self, varChild: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	def accDoDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accHitTest(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> typing.Any:
		...
	def accLocation(self, pxLeft: int=pythoncom.Missing, pyTop: int=pythoncom.Missing, pcxWidth: int=pythoncom.Missing, pcyHeight: int=pythoncom.Missing
			, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accNavigate(self, navDir: int=defaultNamedNotOptArg, varStart: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def accSelect(self, flagsSelect: int=defaultNamedNotOptArg, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
	BeginGroup: typing.Any
	BuiltIn: typing.Any
	BuiltInFace: typing.Any
	Caption: typing.Any
	Control: typing.Any
	Creator: typing.Any
	DescriptionText: typing.Any
	Enabled: typing.Any
	FaceId: typing.Any
	Height: typing.Any
	HelpContextId: typing.Any
	HelpFile: typing.Any
		# Method 'HyperlinkType' returns enumeration of type 'MsoCommandBarButtonHyperlinkType'
	HyperlinkType: MsoCommandBarButtonHyperlinkType
	Id: typing.Any
	Index: typing.Any
	InstanceId: typing.Any
	InstanceIdPtr: typing.Any
	IsPriorityDropped: typing.Any
	Left: typing.Any
		# Method 'Mask' returns object of type 'Picture'
	Mask: Picture
		# Method 'OLEUsage' returns enumeration of type 'MsoControlOLEUsage'
	OLEUsage: MsoControlOLEUsage
	OnAction: typing.Any
	Parameter: typing.Any
		# Method 'Parent' returns object of type 'CommandBar'
	Parent: CommandBar
		# Method 'Picture' returns object of type 'Picture'
	Picture: Picture
	Priority: typing.Any
	ShortcutText: typing.Any
		# Method 'State' returns enumeration of type 'MsoButtonState'
	State: MsoButtonState
		# Method 'Style' returns enumeration of type 'MsoButtonStyle'
	Style: MsoButtonStyle
	Tag: typing.Any
	TooltipText: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'MsoControlType'
	Type: MsoControlType
	Visible: typing.Any
	Width: typing.Any
	accChildCount: typing.Any
	accDefaultAction: typing.Any
	accDescription: typing.Any
	accFocus: typing.Any
	accHelp: typing.Any
	accHelpTopic: typing.Any
	accKeyboardShortcut: typing.Any
	accName: typing.Any
	accParent: typing.Any
	accRole: typing.Any
	accSelection: typing.Any
	accState: typing.Any
	accValue: typing.Any
	def __iter__(self):
		...

class _CommandBarButtonEvents:

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid: typing.Any=defaultNamedNotOptArg, ppvObj: None=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo: int=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, pptinfo: None=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid: typing.Any=defaultNamedNotOptArg, rgszNames: int=defaultNamedNotOptArg, cNames: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg
#			, rgdispid: int=pythoncom.Missing):
#	def OnInvoke(self, dispidMember: int=defaultNamedNotOptArg, riid: typing.Any=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, wFlags: int=defaultNamedNotOptArg
#			, pdispparams: typing.Any=defaultNamedNotOptArg, pvarResult: typing.Any=pythoncom.Missing, pexcepinfo: typing.Any=pythoncom.Missing, puArgErr: int=pythoncom.Missing):
#	def OnClick(self, Ctrl: CommandBarButton=defaultNamedNotOptArg, CancelDefault: bool=defaultNamedNotOptArg):
	...


class _CommandBarComboBox(typing.Protocol):

	def AddItem(self, Text: str=defaultNamedNotOptArg, Index: typing.Any=defaultNamedOptArg) -> None:
		...
	def Clear(self) -> None:
		...
	# Result is of type CommandBarControl
	def Copy(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def Delete(self, Temporary: typing.Any=defaultNamedOptArg) -> None:
		...
	def Execute(self) -> None:
		...
	# The method GetaccDefaultAction is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccDescription is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDescription(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelp is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelp(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelpTopic is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelpTopic(self, pszHelpFile: str=pythoncom.Missing, varChild: typing.Any=defaultNamedOptArg) -> int:
		...
	# The method GetaccKeyboardShortcut is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccKeyboardShortcut(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccName(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccRole is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccRole(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccState is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccState(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccValue(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method List is actually a property, but must be used as a method to correctly pass the arguments
	def List(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	# Result is of type CommandBarControl
	def Move(self, Bar: typing.Any=defaultNamedOptArg, Before: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	def RemoveItem(self, Index: int=defaultNamedNotOptArg) -> None:
		...
	def Reserved1(self) -> None:
		...
	def Reserved2(self) -> None:
		...
	def Reserved3(self) -> None:
		...
	def Reserved4(self) -> None:
		...
	def Reserved5(self) -> None:
		...
	def Reserved6(self) -> None:
		...
	def Reserved7(self) -> None:
		...
	def Reset(self) -> None:
		...
	def SetFocus(self) -> None:
		...
	# The method SetList is actually a property, but must be used as a method to correctly pass the arguments
	def SetList(self, Index: int=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccName(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccValue(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method accChild is actually a property, but must be used as a method to correctly pass the arguments
	def accChild(self, varChild: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	def accDoDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accHitTest(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> typing.Any:
		...
	def accLocation(self, pxLeft: int=pythoncom.Missing, pyTop: int=pythoncom.Missing, pcxWidth: int=pythoncom.Missing, pcyHeight: int=pythoncom.Missing
			, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accNavigate(self, navDir: int=defaultNamedNotOptArg, varStart: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def accSelect(self, flagsSelect: int=defaultNamedNotOptArg, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
	BeginGroup: typing.Any
	BuiltIn: typing.Any
	Caption: typing.Any
	Control: typing.Any
	Creator: typing.Any
	DescriptionText: typing.Any
	DropDownLines: typing.Any
	DropDownWidth: typing.Any
	Enabled: typing.Any
	Height: typing.Any
	HelpContextId: typing.Any
	HelpFile: typing.Any
	Id: typing.Any
	Index: typing.Any
	InstanceId: typing.Any
	InstanceIdPtr: typing.Any
	IsPriorityDropped: typing.Any
	Left: typing.Any
	ListCount: typing.Any
	ListHeaderCount: typing.Any
	ListIndex: typing.Any
		# Method 'OLEUsage' returns enumeration of type 'MsoControlOLEUsage'
	OLEUsage: MsoControlOLEUsage
	OnAction: typing.Any
	Parameter: typing.Any
		# Method 'Parent' returns object of type 'CommandBar'
	Parent: CommandBar
	Priority: typing.Any
		# Method 'Style' returns enumeration of type 'MsoComboStyle'
	Style: MsoComboStyle
	Tag: typing.Any
	Text: typing.Any
	TooltipText: typing.Any
	Top: typing.Any
		# Method 'Type' returns enumeration of type 'MsoControlType'
	Type: MsoControlType
	Visible: typing.Any
	Width: typing.Any
	accChildCount: typing.Any
	accDefaultAction: typing.Any
	accDescription: typing.Any
	accFocus: typing.Any
	accHelp: typing.Any
	accHelpTopic: typing.Any
	accKeyboardShortcut: typing.Any
	accName: typing.Any
	accParent: typing.Any
	accRole: typing.Any
	accSelection: typing.Any
	accState: typing.Any
	accValue: typing.Any
	def __iter__(self):
		...

class _CommandBarComboBoxEvents:

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid: typing.Any=defaultNamedNotOptArg, ppvObj: None=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo: int=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, pptinfo: None=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid: typing.Any=defaultNamedNotOptArg, rgszNames: int=defaultNamedNotOptArg, cNames: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg
#			, rgdispid: int=pythoncom.Missing):
#	def OnInvoke(self, dispidMember: int=defaultNamedNotOptArg, riid: typing.Any=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, wFlags: int=defaultNamedNotOptArg
#			, pdispparams: typing.Any=defaultNamedNotOptArg, pvarResult: typing.Any=pythoncom.Missing, pexcepinfo: typing.Any=pythoncom.Missing, puArgErr: int=pythoncom.Missing):
#	def OnChange(self, Ctrl: CommandBarComboBox=defaultNamedNotOptArg):
	...


class _CommandBars(typing.Protocol):

	# Result is of type CommandBar
	def Add(self, Name: typing.Any=defaultNamedOptArg, Position: typing.Any=defaultNamedOptArg, MenuBar: typing.Any=defaultNamedOptArg, Temporary: typing.Any=defaultNamedOptArg) -> CommandBar:
		...
	# Result is of type CommandBar
	def AddEx(self, TbidOrName: typing.Any=defaultNamedOptArg, Position: typing.Any=defaultNamedOptArg, MenuBar: typing.Any=defaultNamedOptArg, Temporary: typing.Any=defaultNamedOptArg
			, TbtrProtection: typing.Any=defaultNamedOptArg) -> CommandBar:
		...
	def CommitRenderingTransaction(self, hwnd: int=defaultNamedNotOptArg) -> None:
		...
	def ExecuteMso(self, idMso: str=defaultNamedNotOptArg) -> None:
		...
	# Result is of type CommandBarControl
	def FindControl(self, Type: typing.Any=defaultNamedOptArg, Id: typing.Any=defaultNamedOptArg, Tag: typing.Any=defaultNamedOptArg, Visible: typing.Any=defaultNamedOptArg) -> CommandBarControl:
		...
	# Result is of type CommandBarControls
	def FindControls(self, Type: typing.Any=defaultNamedOptArg, Id: typing.Any=defaultNamedOptArg, Tag: typing.Any=defaultNamedOptArg, Visible: typing.Any=defaultNamedOptArg) -> CommandBarControls:
		...
	def GetEnabledMso(self, idMso: str=defaultNamedNotOptArg) -> bool:
		...
	# Result is of type Picture
	def GetImageMso(self, idMso: str=defaultNamedNotOptArg, Width: int=defaultNamedNotOptArg, Height: int=defaultNamedNotOptArg) -> Picture:
		...
	def GetLabelMso(self, idMso: str=defaultNamedNotOptArg) -> str:
		...
	def GetPressedMso(self, idMso: str=defaultNamedNotOptArg) -> bool:
		...
	def GetScreentipMso(self, idMso: str=defaultNamedNotOptArg) -> str:
		...
	def GetSupertipMso(self, idMso: str=defaultNamedNotOptArg) -> str:
		...
	def GetVisibleMso(self, idMso: str=defaultNamedNotOptArg) -> bool:
		...
	# The method IdsString is actually a property, but must be used as a method to correctly pass the arguments
	def IdsString(self, ids: int=defaultNamedNotOptArg, pbstrName: str=pythoncom.Missing) -> int:
		...
	# Result is of type CommandBar
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> CommandBar:
		...
	def ReleaseFocus(self) -> None:
		...
	# The method TmcGetName is actually a property, but must be used as a method to correctly pass the arguments
	def TmcGetName(self, tmc: int=defaultNamedNotOptArg, pbstrName: str=pythoncom.Missing) -> int:
		...
		# Method 'ActionControl' returns object of type 'CommandBarControl'
	ActionControl: CommandBarControl
		# Method 'ActiveMenuBar' returns object of type 'CommandBar'
	ActiveMenuBar: CommandBar
	AdaptiveMenus: typing.Any
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	DisableAskAQuestionDropdown: typing.Any
	DisableCustomize: typing.Any
	DisplayFonts: typing.Any
	DisplayKeysInTooltips: typing.Any
	DisplayTooltips: typing.Any
	LargeButtons: typing.Any
		# Method 'MenuAnimationStyle' returns enumeration of type 'MsoMenuAnimation'
	MenuAnimationStyle: MsoMenuAnimation
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> CommandBar:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class _CommandBarsEvents:

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid: typing.Any=defaultNamedNotOptArg, ppvObj: None=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo: int=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, pptinfo: None=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid: typing.Any=defaultNamedNotOptArg, rgszNames: int=defaultNamedNotOptArg, cNames: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg
#			, rgdispid: int=pythoncom.Missing):
#	def OnInvoke(self, dispidMember: int=defaultNamedNotOptArg, riid: typing.Any=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, wFlags: int=defaultNamedNotOptArg
#			, pdispparams: typing.Any=defaultNamedNotOptArg, pvarResult: typing.Any=pythoncom.Missing, pexcepinfo: typing.Any=pythoncom.Missing, puArgErr: int=pythoncom.Missing):
#	def OnUpdate(self):
	...


class _CustomTaskPane(typing.Protocol):

	def Delete(self) -> None:
		...
	Application: typing.Any
	ContentControl: typing.Any
		# Method 'DockPosition' returns enumeration of type 'MsoCTPDockPosition'
	DockPosition: MsoCTPDockPosition
		# Method 'DockPositionRestrict' returns enumeration of type 'MsoCTPDockPositionRestrict'
	DockPositionRestrict: MsoCTPDockPositionRestrict
	Height: typing.Any
	Title: typing.Any
	Visible: typing.Any
	Width: typing.Any
	Window: typing.Any
	# Default property for this class is 'Title'
	def __call__(self):
		...
	def __iter__(self):
		...

class _CustomTaskPaneEvents:

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid: typing.Any=defaultNamedNotOptArg, ppvObj: None=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo: int=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, pptinfo: None=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid: typing.Any=defaultNamedNotOptArg, rgszNames: int=defaultNamedNotOptArg, cNames: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg
#			, rgdispid: int=pythoncom.Missing):
#	def OnInvoke(self, dispidMember: int=defaultNamedNotOptArg, riid: typing.Any=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, wFlags: int=defaultNamedNotOptArg
#			, pdispparams: typing.Any=defaultNamedNotOptArg, pvarResult: typing.Any=pythoncom.Missing, pexcepinfo: typing.Any=pythoncom.Missing, puArgErr: int=pythoncom.Missing):
#	def OnVisibleStateChange(self, CustomTaskPaneInst: _CustomTaskPane=defaultNamedNotOptArg):
#	def OnDockPositionStateChange(self, CustomTaskPaneInst: _CustomTaskPane=defaultNamedNotOptArg):
	...


class _CustomXMLPart(typing.Protocol):

	def AddNode(self, Parent: CustomXMLNode=defaultNamedNotOptArg, Name: str='', NamespaceURI: str='', NextSibling: CustomXMLNode=0
			, NodeType: MsoCustomXMLNodeType=1, NodeValue: str='') -> None:
		...
	def Delete(self) -> None:
		...
	def Load(self, FilePath: str=defaultNamedNotOptArg) -> bool:
		...
	def LoadXML(self, XML: str=defaultNamedNotOptArg) -> bool:
		...
	# Result is of type CustomXMLNodes
	def SelectNodes(self, XPath: str=defaultNamedNotOptArg) -> CustomXMLNodes:
		...
	# Result is of type CustomXMLNode
	def SelectSingleNode(self, XPath: str=defaultNamedNotOptArg) -> CustomXMLNode:
		...
	Application: typing.Any
	BuiltIn: typing.Any
	Creator: typing.Any
		# Method 'DocumentElement' returns object of type 'CustomXMLNode'
	DocumentElement: CustomXMLNode
		# Method 'Errors' returns object of type 'CustomXMLValidationErrors'
	Errors: CustomXMLValidationErrors
	Id: typing.Any
		# Method 'NamespaceManager' returns object of type 'CustomXMLPrefixMappings'
	NamespaceManager: CustomXMLPrefixMappings
	NamespaceURI: typing.Any
	Parent: typing.Any
		# Method 'SchemaCollection' returns object of type 'CustomXMLSchemaCollection'
	SchemaCollection: CustomXMLSchemaCollection
	XML: typing.Any
	def __iter__(self):
		...

class _CustomXMLPartEvents:

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid: typing.Any=defaultNamedNotOptArg, ppvObj: None=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo: int=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, pptinfo: None=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid: typing.Any=defaultNamedNotOptArg, rgszNames: int=defaultNamedNotOptArg, cNames: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg
#			, rgdispid: int=pythoncom.Missing):
#	def OnInvoke(self, dispidMember: int=defaultNamedNotOptArg, riid: typing.Any=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, wFlags: int=defaultNamedNotOptArg
#			, pdispparams: typing.Any=defaultNamedNotOptArg, pvarResult: typing.Any=pythoncom.Missing, pexcepinfo: typing.Any=pythoncom.Missing, puArgErr: int=pythoncom.Missing):
#	def OnNodeAfterInsert(self, NewNode: CustomXMLNode=defaultNamedNotOptArg, InUndoRedo: bool=defaultNamedNotOptArg):
#	def OnNodeAfterDelete(self, OldNode: CustomXMLNode=defaultNamedNotOptArg, OldParentNode: CustomXMLNode=defaultNamedNotOptArg, OldNextSibling: CustomXMLNode=defaultNamedNotOptArg, InUndoRedo: bool=defaultNamedNotOptArg):
#	def OnNodeAfterReplace(self, OldNode: CustomXMLNode=defaultNamedNotOptArg, NewNode: CustomXMLNode=defaultNamedNotOptArg, InUndoRedo: bool=defaultNamedNotOptArg):
	...


class _CustomXMLParts(typing.Protocol):

	# Result is of type CustomXMLPart
	def Add(self, XML: str='', SchemaCollection: typing.Any=defaultNamedOptArg) -> CustomXMLPart:
		...
	# Result is of type CustomXMLPart
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> CustomXMLPart:
		...
	# Result is of type CustomXMLPart
	def SelectByID(self, Id: str=defaultNamedNotOptArg) -> CustomXMLPart:
		...
	# Result is of type CustomXMLParts
	def SelectByNamespace(self, NamespaceURI: str=defaultNamedNotOptArg) -> CustomXMLParts:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> CustomXMLPart:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class _CustomXMLPartsEvents:

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid: typing.Any=defaultNamedNotOptArg, ppvObj: None=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo: int=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, pptinfo: None=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid: typing.Any=defaultNamedNotOptArg, rgszNames: int=defaultNamedNotOptArg, cNames: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg
#			, rgdispid: int=pythoncom.Missing):
#	def OnInvoke(self, dispidMember: int=defaultNamedNotOptArg, riid: typing.Any=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, wFlags: int=defaultNamedNotOptArg
#			, pdispparams: typing.Any=defaultNamedNotOptArg, pvarResult: typing.Any=pythoncom.Missing, pexcepinfo: typing.Any=pythoncom.Missing, puArgErr: int=pythoncom.Missing):
#	def OnPartAfterAdd(self, NewPart: CustomXMLPart=defaultNamedNotOptArg):
#	def OnPartBeforeDelete(self, OldPart: CustomXMLPart=defaultNamedNotOptArg):
#	def OnPartAfterLoad(self, Part: CustomXMLPart=defaultNamedNotOptArg):
	...


class _CustomXMLSchemaCollection(typing.Protocol):

	# Result is of type CustomXMLSchema
	def Add(self, NamespaceURI: str='', Alias: str='', FileName: str='', InstallForAllUsers: bool=False) -> CustomXMLSchema:
		...
	def AddCollection(self, SchemaCollection: CustomXMLSchemaCollection=defaultNamedNotOptArg) -> None:
		...
	# Result is of type CustomXMLSchema
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index: typing.Any=defaultNamedNotOptArg) -> CustomXMLSchema:
		...
	# The method NamespaceURI is actually a property, but must be used as a method to correctly pass the arguments
	def NamespaceURI(self, Index: int=defaultNamedNotOptArg) -> str:
		...
	def Validate(self) -> bool:
		...
	Application: typing.Any
	Count: typing.Any
	Creator: typing.Any
	Parent: typing.Any
	# Default method for this class is 'Item'
	def __call__(self, Index: typing.Any=defaultNamedNotOptArg) -> CustomXMLSchema:
		...
	def __iter__(self):
		...
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		...
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class _IMsoDispObj(typing.Protocol):

	Application: typing.Any
	Creator: typing.Any
	def __iter__(self):
		...

class _IMsoOleAccDispObj(typing.Protocol):

	# The method GetaccDefaultAction is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccDescription is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccDescription(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelp is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelp(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccHelpTopic is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccHelpTopic(self, pszHelpFile: str=pythoncom.Missing, varChild: typing.Any=defaultNamedOptArg) -> int:
		...
	# The method GetaccKeyboardShortcut is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccKeyboardShortcut(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccName(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method GetaccRole is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccRole(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccState is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccState(self, varChild: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	# The method GetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetaccValue(self, varChild: typing.Any=defaultNamedOptArg) -> str:
		...
	# The method SetaccName is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccName(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method SetaccValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetaccValue(self, varChild: typing.Any=defaultNamedNotOptArg, arg1: str=defaultUnnamedArg) -> None:
		...
	# The method accChild is actually a property, but must be used as a method to correctly pass the arguments
	def accChild(self, varChild: typing.Any=defaultNamedNotOptArg) -> Dispatch:
		...
	def accDoDefaultAction(self, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accHitTest(self, xLeft: int=defaultNamedNotOptArg, yTop: int=defaultNamedNotOptArg) -> typing.Any:
		...
	def accLocation(self, pxLeft: int=pythoncom.Missing, pyTop: int=pythoncom.Missing, pcxWidth: int=pythoncom.Missing, pcyHeight: int=pythoncom.Missing
			, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	def accNavigate(self, navDir: int=defaultNamedNotOptArg, varStart: typing.Any=defaultNamedOptArg) -> typing.Any:
		...
	def accSelect(self, flagsSelect: int=defaultNamedNotOptArg, varChild: typing.Any=defaultNamedOptArg) -> None:
		...
	Application: typing.Any
	Creator: typing.Any
	accChildCount: typing.Any
	accDefaultAction: typing.Any
	accDescription: typing.Any
	accFocus: typing.Any
	accHelp: typing.Any
	accHelpTopic: typing.Any
	accKeyboardShortcut: typing.Any
	accName: typing.Any
	accParent: typing.Any
	accRole: typing.Any
	accSelection: typing.Any
	accState: typing.Any
	accValue: typing.Any
	def __iter__(self):
		...

class _SensitivityLabelEvents:

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid: typing.Any=defaultNamedNotOptArg, ppvObj: None=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo: int=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, pptinfo: None=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid: typing.Any=defaultNamedNotOptArg, rgszNames: int=defaultNamedNotOptArg, cNames: int=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg
#			, rgdispid: int=pythoncom.Missing):
#	def OnInvoke(self, dispidMember: int=defaultNamedNotOptArg, riid: typing.Any=defaultNamedNotOptArg, lcid: int=defaultNamedNotOptArg, wFlags: int=defaultNamedNotOptArg
#			, pdispparams: typing.Any=defaultNamedNotOptArg, pvarResult: typing.Any=pythoncom.Missing, pexcepinfo: typing.Any=pythoncom.Missing, puArgErr: int=pythoncom.Missing):
#	def OnLabelChanged(self, OldLabelInfo: LabelInfo=defaultNamedNotOptArg, NewLabelInfo: LabelInfo=defaultNamedNotOptArg, HResult: int=defaultNamedNotOptArg, Context: Dispatch=defaultNamedNotOptArg):
	...


class CommandBarButton(_CommandBarButton): # A CoClass
	...

class CommandBarComboBox(_CommandBarComboBox): # A CoClass
	...

class CommandBars(_CommandBars): # A CoClass
	...

class CustomTaskPane(_CustomTaskPane): # A CoClass
	...

class CustomXMLPart(_CustomXMLPart): # A CoClass
	...

class CustomXMLParts(_CustomXMLParts): # A CoClass
	...

class CustomXMLSchemaCollection(_CustomXMLSchemaCollection): # A CoClass
	...

class MsoEnvelope(IMsoEnvelopeVB): # A CoClass
	...

class SensitivityLabel(ISensitivityLabel): # A CoClass
	...

