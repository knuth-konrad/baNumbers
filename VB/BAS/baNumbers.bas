Attribute VB_Name = "baNumbers"
'------------------------------------------------------------------------------
'Purpose  : Number helpers - sorting, formatting
'
'Prereq.  : -
'Note     : See https://github.com/knuth-konrad/baNumbers for documentation
'
'   Author: Knuth Konrad 13.03.2015
'   Source: -
'  Changed: 19.04.2016
'           - New functions:
'             baSwapXXX - swaps the contents of two variables
'             TypeOfVariant - determins a VARIANT variables subtype (Integer, Double etc.)
'           28.02.2017
'           - New methods:
'             baRnd: return a pseudo random number (0-1)
'             baRndArray: fill a predim'd array with PRNGs (0-1)
'             baRndRange: return a pseudo random number within a given range
'             baRndRangeArray: fill a predim'd array within a given range
'------------------------------------------------------------------------------
Option Explicit
DefLng A-Z
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
' *** LOCALE definitions
Public Const LANG_NEUTRAL                     As Long = &H0
Public Const LANG_INVARIANT                   As Long = &H7F

Public Const LANG_AFRIKAANS                   As Long = &H36
Public Const LANG_ALBANIAN                    As Long = &H1C
Public Const LANG_ALSATIAN                    As Long = &H84
Public Const LANG_AMHARIC                     As Long = &H5E
Public Const LANG_ARABIC                      As Long = &H1
Public Const LANG_ARMENIAN                    As Long = &H2B
Public Const LANG_ASSAMESE                    As Long = &H4D
Public Const LANG_AZERI                       As Long = &H2C
Public Const LANG_BASHKIR                     As Long = &H6D
Public Const LANG_BASQUE                      As Long = &H2D
Public Const LANG_BELARUSIAN                  As Long = &H23
Public Const LANG_BENGALI                     As Long = &H45
Public Const LANG_BRETON                      As Long = &H7E
Public Const LANG_BOSNIAN                     As Long = &H1A   ' // Use with SUBLANG_BOSNIAN_* Sublanguage IDs
Public Const LANG_BOSNIAN_NEUTRAL             As Long = &H781A '  // Use with the ConvertDefaultLocale function
Public Const LANG_BULGARIAN                   As Long = &H2
Public Const LANG_CATALAN                     As Long = &H3
Public Const LANG_CHINESE                     As Long = &H4
Public Const LANG_CHINESE_SIMPLIFIED          As Long = &H4    ' // Use with the ConvertDefaultLocale function
Public Const LANG_CHINESE_TRADITIONAL         As Long = &H7C04 ' // Use with the ConvertDefaultLocale function
Public Const LANG_CORSICAN                    As Long = &H83
Public Const LANG_CROATIAN                    As Long = &H1A
Public Const LANG_CZECH                       As Long = &H5
Public Const LANG_DANISH                      As Long = &H6
Public Const LANG_DARI                        As Long = &H8C
Public Const LANG_DIVEHI                      As Long = &H65
Public Const LANG_DUTCH                       As Long = &H13
Public Const LANG_ENGLISH                     As Long = &H9
Public Const LANG_ESTONIAN                    As Long = &H25
Public Const LANG_FAEROESE                    As Long = &H38
Public Const LANG_FARSI                       As Long = &H29   ' // Deprecated: use LANG_PERSIAN instead
Public Const LANG_FILIPINO                    As Long = &H64
Public Const LANG_FINNISH                     As Long = &HB
Public Const LANG_FRENCH                      As Long = &HC
Public Const LANG_FRISIAN                     As Long = &H62
Public Const LANG_GALICIAN                    As Long = &H56
Public Const LANG_GEORGIAN                    As Long = &H37
Public Const LANG_GERMAN                      As Long = &H7
Public Const LANG_GREEK                       As Long = &H8
Public Const LANG_GREENLANDIC                 As Long = &H6F
Public Const LANG_GUJARATI                    As Long = &H47
Public Const LANG_HAUSA                       As Long = &H68
Public Const LANG_HEBREW                      As Long = &HD
Public Const LANG_HINDI                       As Long = &H39
Public Const LANG_HUNGARIAN                   As Long = &HE
Public Const LANG_ICELANDIC                   As Long = &HF
Public Const LANG_IGBO                        As Long = &H70
Public Const LANG_INDONESIAN                  As Long = &H21
Public Const LANG_INUKTITUT                   As Long = &H5D
Public Const LANG_IRISH                       As Long = &H3C  ' // Use with the SUBLANG_IRISH_IRELAND Sublanguage ID
Public Const LANG_ITALIAN                     As Long = &H10
Public Const LANG_JAPANESE                    As Long = &H11
Public Const LANG_KANNADA                     As Long = &H4B
Public Const LANG_KASHMIRI                    As Long = &H60
Public Const LANG_KAZAK                       As Long = &H3F
Public Const LANG_KHMER                       As Long = &H53
Public Const LANG_KICHE                       As Long = &H86
Public Const LANG_KINYARWANDA                 As Long = &H87
Public Const LANG_KONKANI                     As Long = &H57
Public Const LANG_KOREAN                      As Long = &H12
Public Const LANG_KYRGYZ                      As Long = &H40
Public Const LANG_LAO                         As Long = &H54
Public Const LANG_LATVIAN                     As Long = &H26
Public Const LANG_LITHUANIAN                  As Long = &H27
Public Const LANG_LOWER_SORBIAN               As Long = &H2E
Public Const LANG_LUXEMBOURGISH               As Long = &H6E
Public Const LANG_MACEDONIAN                  As Long = &H2F   ' the Former Yugoslav Republic of Macedonia
Public Const LANG_MALAY                       As Long = &H3E
Public Const LANG_MALAYALAM                   As Long = &H4C
Public Const LANG_MANIPURI                    As Long = &H58
Public Const LANG_MAORI                       As Long = &H81
Public Const LANG_MAPUDUNGUN                  As Long = &H7A
Public Const LANG_MARATHI                     As Long = &H4E
Public Const LANG_MOHAWK                      As Long = &H7C
Public Const LANG_MONGOLIAN                   As Long = &H50
Public Const LANG_NEPALI                      As Long = &H61
Public Const LANG_NORWEGIAN                   As Long = &H14
Public Const LANG_OCCITAN                     As Long = &H82
Public Const LANG_ORIYA                       As Long = &H48
Public Const LANG_PASHTO                      As Long = &H63
Public Const LANG_PERSIAN                     As Long = &H29
Public Const LANG_POLISH                      As Long = &H15
Public Const LANG_PORTUGUESE                  As Long = &H16
Public Const LANG_PUNJABI                     As Long = &H46
Public Const LANG_QUECHUA                     As Long = &H6B
Public Const LANG_ROMANIAN                    As Long = &H18
Public Const LANG_ROMANSH                     As Long = &H17
Public Const LANG_RUSSIAN                     As Long = &H19
Public Const LANG_SAMI                        As Long = &H3B
Public Const LANG_SANSKRIT                    As Long = &H4F
Public Const LANG_SCOTTISH_GAELIC             As Long = &H91
Public Const LANG_SERBIAN                     As Long = &H1A
Public Const LANG_SERBIAN_NEUTRAL             As Long = &H7C1A   ' // Use with the ConvertDefaultLocale function
Public Const LANG_SINDHI                      As Long = &H59
Public Const LANG_SINHALESE                   As Long = &H5B
Public Const LANG_SLOVAK                      As Long = &H1B
Public Const LANG_SLOVENIAN                   As Long = &H24
Public Const LANG_SOTHO                       As Long = &H6C
Public Const LANG_SPANISH                     As Long = &HA
Public Const LANG_SWAHILI                     As Long = &H41
Public Const LANG_SWEDISH                     As Long = &H1D
Public Const LANG_SYRIAC                      As Long = &H5A
Public Const LANG_TAJIK                       As Long = &H28
Public Const LANG_TAMAZIGHT                   As Long = &H5F
Public Const LANG_TAMIL                       As Long = &H49
Public Const LANG_TATAR                       As Long = &H44
Public Const LANG_TELUGU                      As Long = &H4A
Public Const LANG_THAI                        As Long = &H1E
Public Const LANG_TIBETAN                     As Long = &H51
Public Const LANG_TIGRIGNA                    As Long = &H73
Public Const LANG_TSWANA                      As Long = &H32
Public Const LANG_TURKISH                     As Long = &H1F
Public Const LANG_TURKMEN                     As Long = &H42
Public Const LANG_UIGHUR                      As Long = &H80
Public Const LANG_UKRAINIAN                   As Long = &H22
Public Const LANG_UPPER_SORBIAN               As Long = &H2E
Public Const LANG_URDU                        As Long = &H20
Public Const LANG_UZBEK                       As Long = &H43
Public Const LANG_VIETNAMESE                  As Long = &H2A
Public Const LANG_WELSH                       As Long = &H52
Public Const LANG_WOLOF                       As Long = &H88
Public Const LANG_XHOSA                       As Long = &H34
Public Const LANG_YAKUT                       As Long = &H85
Public Const LANG_YI                          As Long = &H78
Public Const LANG_YORUBA                      As Long = &H6A
Public Const LANG_ZULU                        As Long = &H35

Public Const SUBLANG_NEUTRAL                             As Long = &H0   ' // language neutral
Public Const SUBLANG_DEFAULT                             As Long = &H1   ' // user default
Public Const SUBLANG_SYS_DEFAULT                         As Long = &H2   ' // system default
Public Const SUBLANG_CUSTOM_DEFAULT                      As Long = &H3   ' // default custom language/locale
Public Const SUBLANG_CUSTOM_UNSPECIFIED                  As Long = &H4   ' // custom language/locale
Public Const SUBLANG_UI_CUSTOM_DEFAULT                   As Long = &H5   ' // Default custom MUI language/locale

Public Const SUBLANG_AFRIKAANS_SOUTH_AFRICA              As Long = &H1    ' // Afrikaans (South Africa) As Long = &H0436 af-ZA
Public Const SUBLANG_ALBANIAN_ALBANIA                    As Long = &H1    ' // Albanian (Albania) As Long = &H041c sq-AL
Public Const SUBLANG_ALSATIAN_FRANCE                     As Long = &H1    ' // Alsatian (France) As Long = &H0484
Public Const SUBLANG_AMHARIC_ETHIOPIA                    As Long = &H1    ' // Amharic (Ethiopia) As Long = &H045e
Public Const SUBLANG_ARABIC_SAUDI_ARABIA                 As Long = &H1    ' // Arabic (Saudi Arabia)
Public Const SUBLANG_ARABIC_IRAQ                         As Long = &H2    ' // Arabic (Iraq)
Public Const SUBLANG_ARABIC_EGYPT                        As Long = &H3    ' // Arabic (Egypt)
Public Const SUBLANG_ARABIC_LIBYA                        As Long = &H4    ' // Arabic (Libya)
Public Const SUBLANG_ARABIC_ALGERIA                      As Long = &H5    ' // Arabic (Algeria)
Public Const SUBLANG_ARABIC_MOROCCO                      As Long = &H6    ' // Arabic (Morocco)
Public Const SUBLANG_ARABIC_TUNISIA                      As Long = &H7    ' // Arabic (Tunisia)
Public Const SUBLANG_ARABIC_OMAN                         As Long = &H8    ' // Arabic (Oman)
Public Const SUBLANG_ARABIC_YEMEN                        As Long = &H9    ' // Arabic (Yemen)
Public Const SUBLANG_ARABIC_SYRIA                        As Long = &HA    ' // Arabic (Syria)
Public Const SUBLANG_ARABIC_JORDAN                       As Long = &HB    ' // Arabic (Jordan)
Public Const SUBLANG_ARABIC_LEBANON                      As Long = &HC    ' // Arabic (Lebanon)
Public Const SUBLANG_ARABIC_KUWAIT                       As Long = &HD    ' // Arabic (Kuwait)
Public Const SUBLANG_ARABIC_UAE                          As Long = &HE    ' // Arabic (U.A.E)
Public Const SUBLANG_ARABIC_BAHRAIN                      As Long = &HF    ' // Arabic (Bahrain)
Public Const SUBLANG_ARABIC_QATAR                        As Long = &H10   ' // Arabic (Qatar)
Public Const SUBLANG_ARMENIAN_ARMENIA                    As Long = &H1    ' // Armenian (Armenia) As Long = &H042b hy-AM
Public Const SUBLANG_ASSAMESE_INDIA                      As Long = &H1    ' // Assamese (India) As Long = &H044d
Public Const SUBLANG_AZERI_LATIN                         As Long = &H1    ' // Azeri (Latin)
Public Const SUBLANG_AZERI_CYRILLIC                      As Long = &H2    ' // Azeri (Cyrillic)
Public Const SUBLANG_BASHKIR_RUSSIA                      As Long = &H1    ' // Bashkir (Russia) As Long = &H046d ba-RU
Public Const SUBLANG_BASQUE_BASQUE                       As Long = &H1    ' // Basque (Basque) As Long = &H042d eu-ES
Public Const SUBLANG_BELARUSIAN_BELARUS                  As Long = &H1    ' // Belarusian (Belarus) As Long = &H0423 be-BY
Public Const SUBLANG_BENGALI_INDIA                       As Long = &H1    ' // Bengali (India)
Public Const SUBLANG_BENGALI_BANGLADESH                  As Long = &H2    ' // Bengali (Bangladesh)
Public Const SUBLANG_BOSNIAN_BOSNIA_HERZEGOVINA_LATIN    As Long = &H5    ' // Bosnian (Bosnia and Herzegovina - Latin) As Long = &H141a bs-BA-Latn
Public Const SUBLANG_BOSNIAN_BOSNIA_HERZEGOVINA_CYRILLIC As Long = &H8    ' // Bosnian (Bosnia and Herzegovina - Cyrillic) As Long = &H201a bs-BA-Cyrl
Public Const SUBLANG_BRETON_FRANCE                       As Long = &H1    ' // Breton (France) As Long = &H047e
Public Const SUBLANG_BULGARIAN_BULGARIA                  As Long = &H1    ' // Bulgarian (Bulgaria) As Long = &H0402
Public Const SUBLANG_CATALAN_CATALAN                     As Long = &H1    ' // Catalan (Catalan) As Long = &H0403
Public Const SUBLANG_CHINESE_TRADITIONAL                 As Long = &H1    ' // Chinese (Taiwan Region)
Public Const SUBLANG_CHINESE_SIMPLIFIED                  As Long = &H2    ' // Chinese (PR China)
Public Const SUBLANG_CHINESE_HONGKONG                    As Long = &H3    ' // Chinese (Hong Kong)
Public Const SUBLANG_CHINESE_SINGAPORE                   As Long = &H4    ' // Chinese (Singapore)
Public Const SUBLANG_CHINESE_MACAU                       As Long = &H5    ' // Chinese (Macau)
Public Const SUBLANG_CORSICAN_FRANCE                     As Long = &H1    ' // Corsican (France) As Long = &H0483
Public Const SUBLANG_CZECH_CZECH_REPUBLIC                As Long = &H1    ' // Czech (Czech Republic) As Long = &H0405
Public Const SUBLANG_CROATIAN_CROATIA                    As Long = &H1    ' // Croatian (Croatia)
Public Const SUBLANG_CROATIAN_BOSNIA_HERZEGOVINA_LATIN   As Long = &H4    ' // Croatian (Bosnia and Herzegovina - Latin) As Long = &H101a hr-BA
Public Const SUBLANG_DANISH_DENMARK                      As Long = &H1    ' // Danish (Denmark) As Long = &H0406
Public Const SUBLANG_DARI_AFGHANISTAN                    As Long = &H1    ' // Dari (Afghanistan)
Public Const SUBLANG_DIVEHI_MALDIVES                     As Long = &H1    ' // Divehi (Maldives) As Long = &H0465 div-MV
Public Const SUBLANG_DUTCH                               As Long = &H1    ' // Dutch
Public Const SUBLANG_DUTCH_BELGIAN                       As Long = &H2    ' // Dutch (Belgian)
Public Const SUBLANG_ENGLISH_US                          As Long = &H1    ' // English (USA)
Public Const SUBLANG_ENGLISH_UK                          As Long = &H2    ' // English (UK)
Public Const SUBLANG_ENGLISH_AUS                         As Long = &H3    ' // English (Australian)
Public Const SUBLANG_ENGLISH_CAN                         As Long = &H4    ' // English (Canadian)
Public Const SUBLANG_ENGLISH_NZ                          As Long = &H5    ' // English (New Zealand)
Public Const SUBLANG_ENGLISH_EIRE                        As Long = &H6    ' // English (Irish)
Public Const SUBLANG_ENGLISH_SOUTH_AFRICA                As Long = &H7    ' // English (South Africa)
Public Const SUBLANG_ENGLISH_JAMAICA                     As Long = &H8    ' // English (Jamaica)
Public Const SUBLANG_ENGLISH_CARIBBEAN                   As Long = &H9    ' // English (Caribbean)
Public Const SUBLANG_ENGLISH_BELIZE                      As Long = &HA    ' // English (Belize)
Public Const SUBLANG_ENGLISH_TRINIDAD                    As Long = &HB    ' // English (Trinidad)
Public Const SUBLANG_ENGLISH_ZIMBABWE                    As Long = &HC    ' // English (Zimbabwe)
Public Const SUBLANG_ENGLISH_PHILIPPINES                 As Long = &HD    ' // English (Philippines)
Public Const SUBLANG_ENGLISH_INDIA                       As Long = &H10   ' // English (India)
Public Const SUBLANG_ENGLISH_MALAYSIA                    As Long = &H11   ' // English (Malaysia)
Public Const SUBLANG_ENGLISH_SINGAPORE                   As Long = &H12   ' // English (Singapore)
Public Const SUBLANG_ESTONIAN_ESTONIA                    As Long = &H1    ' // Estonian (Estonia) As Long = &H0425 et-EE
Public Const SUBLANG_FAEROESE_FAROE_ISLANDS              As Long = &H1    ' // Faroese (Faroe Islands) As Long = &H0438 fo-FO
Public Const SUBLANG_FILIPINO_PHILIPPINES                As Long = &H1    ' // Filipino (Philippines) As Long = &H0464 fil-PH
Public Const SUBLANG_FINNISH_FINLAND                     As Long = &H1    ' // Finnish (Finland) As Long = &H040b
Public Const SUBLANG_FRENCH                              As Long = &H1    ' // French
Public Const SUBLANG_FRENCH_BELGIAN                      As Long = &H2    ' // French (Belgian)
Public Const SUBLANG_FRENCH_CANADIAN                     As Long = &H3    ' // French (Canadian)
Public Const SUBLANG_FRENCH_SWISS                        As Long = &H4    ' // French (Swiss)
Public Const SUBLANG_FRENCH_LUXEMBOURG                   As Long = &H5    ' // French (Luxembourg)
Public Const SUBLANG_FRENCH_MONACO                       As Long = &H6    ' // French (Monaco)
Public Const SUBLANG_FRISIAN_NETHERLANDS                 As Long = &H1    ' // Frisian (Netherlands) As Long = &H0462 fy-NL
Public Const SUBLANG_GALICIAN_GALICIAN                   As Long = &H1    ' // Galician (Galician) As Long = &H0456 gl-ES
Public Const SUBLANG_GEORGIAN_GEORGIA                    As Long = &H1    ' // Georgian (Georgia) As Long = &H0437 ka-GE
Public Const SUBLANG_GERMAN                              As Long = &H1    ' // German
Public Const SUBLANG_GERMAN_SWISS                        As Long = &H2    ' // German (Swiss)
Public Const SUBLANG_GERMAN_AUSTRIAN                     As Long = &H3    ' // German (Austrian)
Public Const SUBLANG_GERMAN_LUXEMBOURG                   As Long = &H4    ' // German (Luxembourg)
Public Const SUBLANG_GERMAN_LIECHTENSTEIN                As Long = &H5    ' // German (Liechtenstein)
Public Const SUBLANG_GREEK_GREECE                        As Long = &H1    ' // Greek (Greece)
Public Const SUBLANG_GREENLANDIC_GREENLAND               As Long = &H1    ' // Greenlandic (Greenland) As Long = &H046f kl-GL
Public Const SUBLANG_GUJARATI_INDIA                      As Long = &H1    ' // Gujarati (India (Gujarati Script)) As Long = &H0447 gu-IN
Public Const SUBLANG_HAUSA_NIGERIA_LATIN                 As Long = &H1    ' // Hausa (Latin, Nigeria) As Long = &H0468 ha-NG-Latn
Public Const SUBLANG_HEBREW_ISRAEL                       As Long = &H1    ' // Hebrew (Israel) As Long = &H040d
Public Const SUBLANG_HINDI_INDIA                         As Long = &H1    ' // Hindi (India) As Long = &H0439 hi-IN
Public Const SUBLANG_HUNGARIAN_HUNGARY                   As Long = &H1    ' // Hungarian (Hungary) As Long = &H040e
Public Const SUBLANG_ICELANDIC_ICELAND                   As Long = &H1    ' // Icelandic (Iceland) As Long = &H040f
Public Const SUBLANG_IGBO_NIGERIA                        As Long = &H1    ' // Igbo (Nigeria) As Long = &H0470 ig-NG
Public Const SUBLANG_INDONESIAN_INDONESIA                As Long = &H1    ' // Indonesian (Indonesia) As Long = &H0421 id-ID
Public Const SUBLANG_INUKTITUT_CANADA                    As Long = &H1    ' // Inuktitut (Syllabics) (Canada) As Long = &H045d iu-CA-Cans
Public Const SUBLANG_INUKTITUT_CANADA_LATIN              As Long = &H2    ' // Inuktitut (Canada - Latin)
Public Const SUBLANG_IRISH_IRELAND                       As Long = &H2    ' // Irish (Ireland)
Public Const SUBLANG_ITALIAN                             As Long = &H1    ' // Italian
Public Const SUBLANG_ITALIAN_SWISS                       As Long = &H2    ' // Italian (Swiss)
Public Const SUBLANG_JAPANESE_JAPAN                      As Long = &H1    ' // Japanese (Japan) As Long = &H0411
Public Const SUBLANG_KANNADA_INDIA                       As Long = &H1    ' // Kannada (India (Kannada Script)) As Long = &H044b kn-IN
Public Const SUBLANG_KASHMIRI_SASIA                      As Long = &H2    ' // Kashmiri (South Asia)
Public Const SUBLANG_KASHMIRI_INDIA                      As Long = &H2    ' // Kashmiri (India)
' //Public Const SUBLANG_KASHMIRI_SASIA                      As Long = &H2    ' // Kashmiri (South Asia)
' //Public Const SUBLANG_KASHMIRI_INDIA                      As Long = &H2    ' // For app compatibility only
Public Const SUBLANG_KAZAK_KAZAKHSTAN                    As Long = &H1    ' // Kazakh (Kazakhstan) As Long = &H043f kk-KZ
Public Const SUBLANG_KHMER_CAMBODIA                      As Long = &H1    ' // Khmer (Cambodia) As Long = &H0453 kh-KH
Public Const SUBLANG_KICHE_GUATEMALA                     As Long = &H1    ' // K' //iche (Guatemala)
Public Const SUBLANG_KINYARWANDA_RWANDA                  As Long = &H1    ' // Kinyarwanda (Rwanda) As Long = &H0487 rw-RW
Public Const SUBLANG_KONKANI_INDIA                       As Long = &H1    ' // Konkani (India) As Long = &H0457 kok-IN
Public Const SUBLANG_KOREAN                              As Long = &H1    ' // Korean (Extended Wansung)
Public Const SUBLANG_KYRGYZ_KYRGYZSTAN                   As Long = &H1    ' // Kyrgyz (Kyrgyzstan) As Long = &H0440 ky-KG
Public Const SUBLANG_LAO_LAO                             As Long = &H1    ' // Lao (Lao PDR) As Long = &H0454 lo-LA
Public Const SUBLANG_LATVIAN_LATVIA                      As Long = &H1    ' // Latvian (Latvia) As Long = &H0426 lv-LV
Public Const SUBLANG_LITHUANIAN                          As Long = &H1    ' // Lithuanian
' //Public Const SUBLANG_LITHUANIAN_CLASSIC       As Long = &H02  ' // Lithuanian (Classic)  Not in Windows 7 SDK
Public Const SUBLANG_LOWER_SORBIAN_GERMANY               As Long = &H2    ' // Lower Sorbian (Germany) As Long = &H082e wee-DE
Public Const SUBLANG_LUXEMBOURGISH_LUXEMBOURG            As Long = &H1    ' // Luxembourgish (Luxembourg) As Long = &H046e lb-LU
Public Const SUBLANG_MACEDONIAN_MACEDONIA                As Long = &H1    ' // Macedonian (Macedonia (FYROM)) As Long = &H042f mk-MK
Public Const SUBLANG_MALAY_MALAYSIA                      As Long = &H1    ' // Malay (Malaysia)
Public Const SUBLANG_MALAY_BRUNEI_DARUSSALAM             As Long = &H2    ' // Malay (Brunei Darussalam)
Public Const SUBLANG_MALAYALAM_INDIA                     As Long = &H1    ' // Malayalam (India (Malayalam Script) ) As Long = &H044c ml-IN
Public Const SUBLANG_MALTESE_MALTA                       As Long = &H1    ' // Maltese (Malta) As Long = &H043a mt-MT
Public Const SUBLANG_MAORI_NEW_ZEALAND                   As Long = &H1    ' // Maori (New Zealand) As Long = &H0481 mi-NZ
Public Const SUBLANG_MAPUDUNGUN_CHILE                    As Long = &H1    ' // Mapudungun (Chile) As Long = &H047a arn-CL
Public Const SUBLANG_MARATHI_INDIA                       As Long = &H1    ' // Marathi (India) As Long = &H044e mr-IN
Public Const SUBLANG_MOHAWK_MOHAWK                       As Long = &H1    ' // Mohawk (Mohawk) As Long = &H047c moh-CA
Public Const SUBLANG_MONGOLIAN_CYRILLIC_MONGOLIA         As Long = &H1    ' // Mongolian (Cyrillic, Mongolia)
Public Const SUBLANG_MONGOLIAN_PRC                       As Long = &H2    ' // Mongolian (PRC)
Public Const SUBLANG_NEPALI_INDIA                        As Long = &H2    ' // Nepali (India)
Public Const SUBLANG_NEPALI_NEPAL                        As Long = &H1    ' // Nepali (Nepal) As Long = &H0461 ne-NP
Public Const SUBLANG_NORWEGIAN_BOKMAL                    As Long = &H1    ' // Norwegian (Bokmal)
Public Const SUBLANG_NORWEGIAN_NYNORSK                   As Long = &H2    ' // Norwegian (Nynorsk)
Public Const SUBLANG_OCCITAN_FRANCE                      As Long = &H1    ' // Occitan (France) As Long = &H0482 oc-FR
Public Const SUBLANG_ORIYA_INDIA                         As Long = &H1    ' // Oriya (India (Oriya Script)) As Long = &H0448 or-IN
Public Const SUBLANG_PASHTO_AFGHANISTAN                  As Long = &H1    ' // Pashto (Afghanistan)
Public Const SUBLANG_PERSIAN_IRAN                        As Long = &H1    ' // Persian (Iran) As Long = &H0429 fa-IR
Public Const SUBLANG_POLISH_POLAND                       As Long = &H1    ' // Polish (Poland) As Long = &H0415
Public Const SUBLANG_PORTUGUESE                          As Long = &H2    ' // Portuguese
Public Const SUBLANG_PORTUGUESE_BRAZILIAN                As Long = &H1    ' // Portuguese (Brazilian)
Public Const SUBLANG_PUNJABI_INDIA                       As Long = &H1    ' // Punjabi (India (Gurmukhi Script)) As Long = &H0446 pa-IN
Public Const SUBLANG_QUECHUA_BOLIVIA                     As Long = &H1    ' // Quechua (Bolivia)
Public Const SUBLANG_QUECHUA_ECUADOR                     As Long = &H2    ' // Quechua (Ecuador)
Public Const SUBLANG_QUECHUA_PERU                        As Long = &H3    ' // Quechua (Peru)
Public Const SUBLANG_ROMANIAN_ROMANIA                    As Long = &H1    ' // Romanian (Romania) As Long = &H0418
Public Const SUBLANG_ROMANSH_SWITZERLAND                 As Long = &H1    ' // Romansh (Switzerland) As Long = &H0417 rm-CH
Public Const SUBLANG_RUSSIAN_RUSSIA                      As Long = &H1    ' // Russian (Russia) As Long = &H0419
Public Const SUBLANG_SAMI_NORTHERN_NORWAY                As Long = &H1    ' // Northern Sami (Norway)
Public Const SUBLANG_SAMI_NORTHERN_SWEDEN                As Long = &H2    ' // Northern Sami (Sweden)
Public Const SUBLANG_SAMI_NORTHERN_FINLAND               As Long = &H3    ' // Northern Sami (Finland)
Public Const SUBLANG_SAMI_LULE_NORWAY                    As Long = &H4    ' // Lule Sami (Norway)
Public Const SUBLANG_SAMI_LULE_SWEDEN                    As Long = &H5    ' // Lule Sami (Sweden)
Public Const SUBLANG_SAMI_SOUTHERN_NORWAY                As Long = &H6    ' // Southern Sami (Norway)
Public Const SUBLANG_SAMI_SOUTHERN_SWEDEN                As Long = &H7    ' // Southern Sami (Sweden)
Public Const SUBLANG_SAMI_SKOLT_FINLAND                  As Long = &H8    ' // Skolt Sami (Finland)
Public Const SUBLANG_SAMI_INARI_FINLAND                  As Long = &H9    ' // Inari Sami (Finland)
Public Const SUBLANG_SANSKRIT_INDIA                      As Long = &H1    ' // Sanskrit (India) As Long = &H044f sa-IN
Public Const SUBLANG_SCOTTISH_GAELIC                     As Long = &H1    ' // Scottish Gaelic (United Kingdom) 0x0491 gd-GB
Public Const SUBLANG_SERBIAN_BOSNIA_HERZEGOVINA_LATIN    As Long = &H6    ' // Serbian (Bosnia and Herzegovina - Latin)
Public Const SUBLANG_SERBIAN_BOSNIA_HERZEGOVINA_CYRILLIC As Long = &H7    ' // Serbian (Bosnia and Herzegovina - Cyrillic)
Public Const SUBLANG_SERBIAN_MONTENEGRO_LATIN            As Long = &HB    ' // Serbian (Montenegro - Latn)
Public Const SUBLANG_SERBIAN_MONTENEGRO_CYRILLIC         As Long = &HC    ' // Serbian (Montenegro - Cyrillic)
Public Const SUBLANG_SERBIAN_SERBIA_LATIN                As Long = &H9    ' // Serbian (Serbia - Latin)
Public Const SUBLANG_SERBIAN_SERBIA_CYRILLIC             As Long = &HA    ' // Serbian (Serbia - Cyrillic)
Public Const SUBLANG_SERBIAN_CROATIA                     As Long = &H1    ' // Croatian (Croatia) As Long = &H041a hr-HR
Public Const SUBLANG_SERBIAN_LATIN                       As Long = &H2    ' // Serbian (Latin)
Public Const SUBLANG_SERBIAN_CYRILLIC                    As Long = &H3    ' // Serbian (Cyrillic)
Public Const SUBLANG_SINDHI_INDIA                        As Long = &H1    ' // Sindhi (India) reserved As Long = &H0459
Public Const SUBLANG_SINDHI_PAKISTAN                     As Long = &H2    ' // Sindhi (Pakistan) reserved As Long = &H0859
Public Const SUBLANG_SINDHI_AFGHANISTAN                  As Long = &H2    ' // For app compatibility only
Public Const SUBLANG_SINHALESE_SRI_LANKA                 As Long = &H1    ' // Sinhalese (Sri Lanka)
Public Const SUBLANG_SOTHO_NORTHERN_SOUTH_AFRICA         As Long = &H1    ' // Northern Sotho (South Africa)
Public Const SUBLANG_SLOVAK_SLOVAKIA                     As Long = &H1    ' // Slovak (Slovakia) As Long = &H041b sk-SK
Public Const SUBLANG_SLOVENIAN_SLOVENIA                  As Long = &H1    ' // Slovenian (Slovenia) As Long = &H0424 sl-SI
Public Const SUBLANG_SPANISH                             As Long = &H1    ' // Spanish (Castilian)
Public Const SUBLANG_SPANISH_MEXICAN                     As Long = &H2    ' // Spanish (Mexican)
Public Const SUBLANG_SPANISH_MODERN                      As Long = &H3    ' // Spanish (Modern)
Public Const SUBLANG_SPANISH_GUATEMALA                   As Long = &H4    ' // Spanish (Guatemala)
Public Const SUBLANG_SPANISH_COSTA_RICA                  As Long = &H5    ' // Spanish (Costa Rica)
Public Const SUBLANG_SPANISH_PANAMA                      As Long = &H6    ' // Spanish (Panama)
Public Const SUBLANG_SPANISH_DOMINICAN_REPUBLIC          As Long = &H7    ' // Spanish (Dominican Republic)
Public Const SUBLANG_SPANISH_VENEZUELA                   As Long = &H8    ' // Spanish (Venezuela)
Public Const SUBLANG_SPANISH_COLOMBIA                    As Long = &H9    ' // Spanish (Colombia)
Public Const SUBLANG_SPANISH_PERU                        As Long = &HA    ' // Spanish (Peru)
Public Const SUBLANG_SPANISH_ARGENTINA                   As Long = &HB    ' // Spanish (Argentina)
Public Const SUBLANG_SPANISH_ECUADOR                     As Long = &HC    ' // Spanish (Ecuador)
Public Const SUBLANG_SPANISH_CHILE                       As Long = &HD    ' // Spanish (Chile)
Public Const SUBLANG_SPANISH_URUGUAY                     As Long = &HE    ' // Spanish (Uruguay)
Public Const SUBLANG_SPANISH_PARAGUAY                    As Long = &HF    ' // Spanish (Paraguay)
Public Const SUBLANG_SPANISH_BOLIVIA                     As Long = &H10   ' // Spanish (Bolivia)
Public Const SUBLANG_SPANISH_EL_SALVADOR                 As Long = &H11   ' // Spanish (El Salvador)
Public Const SUBLANG_SPANISH_HONDURAS                    As Long = &H12   ' // Spanish (Honduras)
Public Const SUBLANG_SPANISH_NICARAGUA                   As Long = &H13   ' // Spanish (Nicaragua)
Public Const SUBLANG_SPANISH_PUERTO_RICO                 As Long = &H14   ' // Spanish (Puerto Rico)
Public Const SUBLANG_SPANISH_US                          As Long = &H15   ' // Spanish (United States)
Public Const SUBLANG_SWAHILI_KENYA                       As Long = &H1    ' // Swahili (Kenya) As Long = &H0441 sw-KE
Public Const SUBLANG_SWEDISH                             As Long = &H1    ' // Swedish
Public Const SUBLANG_SWEDISH_FINLAND                     As Long = &H2    ' // Swedish (Finland)
Public Const SUBLANG_SYRIAC_SYRIA                        As Long = &H1    ' // Syriac (Syria) As Long = &H045a syr-SY
Public Const SUBLANG_TAJIK_TAJIKISTAN                    As Long = &H1    ' // Tajik (Tajikistan) As Long = &H0428 tg-TJ-Cyrl
Public Const SUBLANG_TAMAZIGHT_ALGERIA_LATIN             As Long = &H2    ' // Tamazight (Latin, Algeria) As Long = &H085f tmz-DZ-Latn
Public Const SUBLANG_TAMIL_INDIA                         As Long = &H1    ' // Tamil (India)
Public Const SUBLANG_TATAR_RUSSIA                        As Long = &H1    ' // Tatar (Russia) As Long = &H0444 tt-RU
Public Const SUBLANG_TELUGU_INDIA                        As Long = &H1    ' // Telugu (India (Telugu Script)) As Long = &H044a te-IN
Public Const SUBLANG_THAI_THAILAND                       As Long = &H1    ' // Thai (Thailand) As Long = &H041e th-TH
Public Const SUBLANG_TIBETAN_PRC                         As Long = &H1    ' // Tibetan (PRC)
Public Const SUBLANG_TIGRIGNA_ERITREA                    As Long = &H2    ' // Tigrigna (Eritrea)
Public Const SUBLANG_TSWANA_SOUTH_AFRICA                 As Long = &H1    ' // Setswana / Tswana (South Africa) As Long = &H0432 tn-ZA
Public Const SUBLANG_TURKISH_TURKEY                      As Long = &H1    ' // Turkish (Turkey) As Long = &H041f tr-TR
Public Const SUBLANG_TURKMEN_TURKMENISTAN                As Long = &H1    ' // Turkmen (Turkmenistan) As Long = &H0442 tk-TM
Public Const SUBLANG_UIGHUR_PRC                          As Long = &H1    ' // Uighur (PRC) As Long = &H0480 ug-CN
Public Const SUBLANG_UKRAINIAN_UKRAINE                   As Long = &H1    ' // Ukrainian (Ukraine) As Long = &H0422 uk-UA
Public Const SUBLANG_UPPER_SORBIAN_GERMANY               As Long = &H1    ' // Upper Sorbian (Germany) As Long = &H042e wen-DE
Public Const SUBLANG_URDU_PAKISTAN                       As Long = &H1    ' // Urdu (Pakistan)
Public Const SUBLANG_URDU_INDIA                          As Long = &H2    ' // Urdu (India)
Public Const SUBLANG_UZBEK_LATIN                         As Long = &H1    ' // Uzbek (Latin)
Public Const SUBLANG_UZBEK_CYRILLIC                      As Long = &H2    ' // Uzbek (Cyrillic)
Public Const SUBLANG_VIETNAMESE_VIETNAM                  As Long = &H1    ' // Vietnamese (Vietnam) As Long = &H042a vi-VN
Public Const SUBLANG_WELSH_UNITED_KINGDOM                As Long = &H1    ' // Welsh (United Kingdom) As Long = &H0452 cy-GB
Public Const SUBLANG_WOLOF_SENEGAL                       As Long = &H1    ' // Wolof (Senegal)
Public Const SUBLANG_XHOSA_SOUTH_AFRICA                  As Long = &H1    ' // isiXhosa / Xhosa (South Africa) As Long = &H0434 xh-ZA
Public Const SUBLANG_YAKUT_RUSSIA                        As Long = &H1    ' // Yakut (Russia) As Long = &H0485 sah-RU
Public Const SUBLANG_YI_PRC                              As Long = &H1    ' // Yi (PRC)) As Long = &H0478
Public Const SUBLANG_YORUBA_NIGERIA                      As Long = &H1    ' // Yoruba (Nigeria) 046a yo-NG
Public Const SUBLANG_ZULU_SOUTH_AFRICA                   As Long = &H1    ' // isiZulu / Zulu (South Africa) As Long = &H0435 zu-ZA
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
Public Type NUMBERFMT
   NumDigits As Long
   LeadingZero As Long
   Grouping As Long
   lpDecimalSep As String
   lpThousandSep As String
   NegativeOrder As Long
End Type

' Enum for LocaleString
Public Enum eLocaleStringConstants
  locGrouping = &H10
  locNumDigits = &H11
  locLeadingZero = &H12
  locCurrency = &H14
  locCurSymbol = &H15
  locDate = &H1D
  locDecimal = &HE
  locList = &HC
  locMoneyDecimal = &H16
  locMoneyThousands = &H17
  locNegative = &H51
  locPositive = &H50
  locThousands = &HF
  locTime = &H1E
End Enum

' Enum for TypeOfVariant
Public Enum eVariantType
   vtErr = -1                    ' User-defined variable to pass processing error back to caller
                                 ' Don't mistake it as vtError (%VT_ERROR) below
   vtEmpty = 0                   ' %VT_EMPTY
   vtNull = 1                    ' %VT_NULL
   vtIntegerSigned = 2           ' %VT_I2
   vtLongIntegerSigned = 3       ' %VT_I4
   vtSingle = 4                  ' %VT_R4
   vtDouble = 5                  ' %VT_R8
   vtCurrency = 6                ' %VT_CY
   vtDate = 7                    ' %VT_DATE
   vtString = 8                  ' %VT_BSTR
   vtIDispatch = 9               ' %VT_IDISPATCH (Object with IDispatch interface)
   vtError = 10                  ' %VT_ERROR
   vtBool = 11                   ' %VT_BOOL (basically same as vtByteUnsigned, see below)
   vtVariant = 12                ' %VT_VARIANT
   vtIUnknown = 13               ' %VT_UNKNOWN (Object with IUnknown interface)
   vtDecimal = 14                ' %VT_DECIMAL
   vtByteSigned = 16             ' %VT_I1
   vtByteUnsigned = 17           ' %VT_UI1
   vtIntegerUnsigned = 18        ' %VT_UI2
   vtLongIntegerUnsigned = 19    ' %VT_UI4
   vtQuadIntegerSigned = 20      ' %VT_I8
   vtQuadIntegerUnsigned = 21    ' %VT_UI8
   vtLongInteger = 22            ' %VT_INT (basically same as vtLongIntegerSigned)
   vtDWord = 23                  ' %VT_UINT (basically same as vtLongIntegerUnsigned)
   vtVoid = 24                   ' %VT_VOID (A C-style void type)
   vtHResult = 25                ' %VT_HRESULT (COM result code)
   vtPointer = 26                ' %VT_PTR
   vtSafeArray = 27              ' %VT_SAFEARRAY (a VB array)
   vtCArray = 28                 ' %VT_CARRAY (A C-style array)
   vtUDT = 29                    ' %VT_USERDEFINED
   vtLPStr = 30                  ' %VT_LPSTR (ANSI string)
   vtLPWStr = 31                 ' %VT_LPWSTR (Unicode string)
   vtFileTime = 64               ' %VT_FILETIME (Win API FILETIME)
   vtBlob = 65                   ' %VT_BLOB (An arbitrary block of memory)
   vtStream = 66                 ' %VT_STREAM (A stream of bytes)
   vtStorage = 67                ' %VT_STORAGE
   vtStreamedObject = 68         ' %VT_STREAMED_OBJECT (A stream that contains an object)
   vtStoredObject = 69           ' %VT_STORED_OBJECT (A storage object)
   vtBlobObject = 70             ' %VT_BLOB_OBJECT (A block of memory that represents an object)
   vtClipboardFormat = 71        ' %VT_CF (Clipboard format)
   vtCLSID = 72                  ' %VT_CLSID (Class ID)
   vtVector = &H1000             ' %VT_VECTOR (An array with a leading count)
   vtArray = &H2000              ' %VT_ARRAY
   vtByRef = &H4000              ' %VT_BYREF (A reference value)
End Enum

' Array data type identifier used by baSort2Arrays()
Public Enum eArrayDataType
   adtByte = 0
   adtCurrency = 1
   adtDouble = 2
   adtInteger = 3
   adtLong = 4
   adtSingle = 5
   adtString = 6
End Enum
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
' *** Win32 API
Private Declare Function GetNumberFormat Lib "kernel32" Alias "GetNumberFormatA" (ByVal Locale As Long, _
   ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As NUMBERFMT, ByVal lpNumberStr _
   As String, ByVal cchNumber As Long) As Long

' *** baNumbers.dll
' Pass the array like
' Dim a() As Integer : baSortInteger a()
Declare Sub baSortByte Lib "baNumbers.dll" (a() As Byte, Optional ByVal bolDescending As Boolean = False)
Declare Sub baSortCurrency Lib "baNumbers.dll" (a() As Currency, Optional ByVal bolDescending As Boolean = False)
Declare Sub baSortDouble Lib "baNumbers.dll" (a() As Double, Optional ByVal bolDescending As Boolean = False)
Declare Sub baSortInteger Lib "baNumbers.dll" (a() As Integer, Optional ByVal bolDescending As Boolean = False)
Declare Sub baSortLong Lib "baNumbers.dll" (a() As Long, Optional ByVal bolDescending As Boolean = False)
Declare Sub baSortSingle Lib "baNumbers.dll" (a() As Single, Optional ByVal bolDescending As Boolean = False)

Declare Function baMedianByte Lib "baNumbers.dll" (a() As Byte) As Byte
Declare Function baMedianCurrency Lib "baNumbers.dll" (a() As Currency) As Currency
Declare Function baMedianDouble Lib "baNumbers.dll" (a() As Double) As Double
Declare Function baMedianInteger Lib "baNumbers.dll" (a() As Integer) As Integer
Declare Function baMedianLong Lib "baNumbers.dll" (a() As Long) As Long
Declare Function baMedianSingle Lib "baNumbers.dll" (a() As Single) As Single

Declare Function baSetByte Lib "baNumbers.dll" (a() As Byte, ByVal value As Byte) As Boolean
Declare Function baSetInteger Lib "baNumbers.dll" (a() As Integer, ByVal value As Integer) As Boolean
Declare Function baSetLong Lib "baNumbers.dll" (a() As Long, ByVal value As Long) As Boolean
Declare Function baSetSingle Lib "baNumbers.dll" (a() As Single, ByVal value As Single) As Boolean
Declare Function baSetDouble Lib "baNumbers.dll" (a() As Double, ByVal value As Double) As Boolean

Declare Function baFormatNumber Lib "baNumbers.dll" (ByVal curNumber As Currency, _
   ByVal wLangLocale As Long, ByVal wSubLangLocale As Long) As String
Declare Function baFormatNumberEx Lib "baNumbers.dll" (ByVal curNumber As Currency, _
   ByVal dwLangID As Long) As String

Declare Function baFracCur Lib "baNumbers.dll" (ByVal curValue As Currency) As Currency
Declare Function baFracDouble Lib "baNumbers.dll" (ByVal dblValue As Double) As Double
Declare Function baFracSingle Lib "baNumbers.dll" (ByVal fValue As Single) As Single

Declare Function LCIDFromLangID Lib "baNumbers.dll" (ByVal dwLangID As Long) As Long
Declare Function LocaleString Lib "baNumbers.dll" (ByVal dwLCID As Long, ByVal eInfo As eLocaleStringConstants) As String

Declare Function baSwapByte Lib "baNumbers.dll" (ByRef v1 As Byte, ByRef v2 As Byte) As Boolean
Declare Function baSwapCurrency Lib "baNumbers.dll" (ByRef v1 As Currency, ByRef v2 As Currency) As Boolean
Declare Function baSwapDouble Lib "baNumbers.dll" (ByRef v1 As Double, ByRef v2 As Double) As Boolean
Declare Function baSwapInteger Lib "baNumbers.dll" (ByRef v1 As Integer, ByRef v2 As Integer) As Boolean
Declare Function baSwapLong Lib "baNumbers.dll" (ByRef v1 As Long, ByRef v2 As Long) As Boolean
Declare Function baSwapSingle Lib "baNumbers.dll" (ByRef v1 As Single, ByRef v2 As Single) As Boolean

Declare Function TypeOfVariant Lib "baNumbers.dll" (ByVal vnt As Variant) As eVariantType

Declare Function Int2Wrd Lib "baNumbers.dll" (ByVal iValue As Integer) As Long
Declare Function Int2DWrd Lib "baNumbers.dll" (ByVal iValue As Integer) As Currency
Declare Function Lng2DWrd Lib "baNumbers.dll" (ByVal lValue As Long) As Currency
Declare Function Lng2Quad Lib "baNumbers.dll" (ByVal lValue As Long) As Currency

Declare Function baRnd Lib "baNumbers.dll" () As Currency
Declare Sub baRndArray Lib "baNumbers.dll" (a() As Currency)
Declare Function baRndRange Lib "baNumbers.dll" (ByVal lLower As Long, ByVal lUpper As Long) As Long
Declare Sub baRndRangeArrayLong Lib "baNumbers.dll" (a() As Long, ByVal lLower As Long, ByVal lUpper As Long)

' *** Helper methods for baSort2Arrays()
' Each of the following methods is used to set 1 of 2 arrays that then
' will be used with method baSort2Arrays(). Your code must therefore consist
' of two baArray<DataType>Set lines *before* you call baSort2Arrays()
' Both array also most have the same number of dimensions and elements.
Declare Function baArrayByteSet Lib "baNumbers.dll" (a() As Byte, ByVal WhichArray As Long) As Boolean
Declare Function baArrayCurrencySet Lib "baNumbers.dll" (a() As Currency, ByVal WhichArray As Long) As Boolean
Declare Function baArrayDoubleSet Lib "baNumbers.dll" (a() As Double, ByVal WhichArray As Long) As Boolean
Declare Function baArrayIntegerSet Lib "baNumbers.dll" (a() As Integer, ByVal WhichArray As Long) As Boolean
Declare Function baArraySingleSet Lib "baNumbers.dll" (a() As Single, ByVal WhichArray As Long) As Boolean
Declare Function baArrayStringSet Lib "baNumbers.dll" (a() As String, ByVal WhichArray As Long) As Boolean

Declare Function baSort2Arrays Lib "baNumbers.dll" (ByVal eDataType1 As eArrayDataType, _
   eDataType2 As eArrayDataType, Optional ByVal bolDescending As Boolean = False) As Long

'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
'==============================================================================

Public Function baFormatNumberForLCID(ByVal curValue As Currency, ByVal dwLCID As Long, _
   Optional ByVal lNumFractionalDigits As Long = -1, Optional ByVal bolAddLeadingZero As Boolean = True) As String
'------------------------------------------------------------------------------
'Purpose  : Return properly formatted string representation for a number, according to
'           LCIDs setting.
'
'Prereq.  : -
'Parameter: curValue             - Number to format
'           wLCID                - Locale for which to format the number
'           lNumFractionalDigits - Format with <n> fractional digits, -1 = use default for LCID
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 13.03.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim udtNF As NUMBERFMT
   Dim sValue As String
   Dim sRet As String, lRet As Long
   
   Dim lpzInputValue As String      '18 digits, leading zero, optional leading minus, and decimal point.
   
   ' Stringbuffer initialisieren
   sValue = Space$(256)
   
   With udtNF
   
      ' Grouping
      sRet = LocaleString(dwLCID, locGrouping)
      .Grouping = Val(sRet)
   
      ' NumDigits
      If lNumFractionalDigits = -1 Then
      ' Standard für LCID verwenden
         sRet = LocaleString(dwLCID, locNumDigits)
         lNumFractionalDigits = Val(sRet)
      End If
      .NumDigits = lNumFractionalDigits
   
      ' LeadingZero
      If bolAddLeadingZero = True Then
         sRet = LocaleString(dwLCID, locLeadingZero)
         .LeadingZero = Val(sRet)
      End If
      
      ' lpDecimalSep
      .lpDecimalSep = LocaleString(dwLCID, locDecimal)
      
      ' lpThousandSep
      .lpThousandSep = LocaleString(dwLCID, locThousands)
      
   End With
   
   ' lpzInputValue = LTrim$(Str$(curNumber, 10))
   lpzInputValue = LTrim$(Str$(curValue))
   
   lRet = GetNumberFormat(dwLCID, 0, lpzInputValue, udtNF, sValue, Len(sValue))
   If lRet > 0 Then
      sRet = Left$(sValue, lRet)
      baFormatNumberForLCID = Trim0(sRet)
   End If
   
End Function
'==============================================================================

Public Function baTypeOfVariantToString(ByVal eValue As eVariantType) As String
'------------------------------------------------------------------------------
'Purpose  : Return the string representation of the variant subtype
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 13.05.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim sResult As String
   
   Select Case eValue
   
   Case vtErr
      sResult = "vtErr"
   Case vtEmpty
      sResult = "vtEmpty"
   Case vtNull
      sResult = "vtNull"
   Case vtIntegerSigned
      sResult = "vtIntegerSigned"
   Case vtLongIntegerSigned
      sResult = "vtLongIntegerSigned"
   Case vtSingle
      sResult = "vtSingle"
   Case vtDouble
      sResult = "vtDouble"
   Case vtCurrency
      sResult = "vtCurrency"
   Case vtDate
      sResult = "vtDate"
   Case vtString
      sResult = "vtString"
   Case vtIDispatch
      sResult = "vtIDispatch"
   Case vtError
      sResult = "vtError"
   Case vtBool
      sResult = "vtBool"
   Case vtVariant
      sResult = "vtVariant"
   Case vtIUnknown
      sResult = "vtIUnknown"
   Case vtDecimal
      sResult = "vtDecimal"
   Case vtByteSigned
      sResult = "vtByteSigned"
   Case vtByteUnsigned
      sResult = "vtByteUnsigned"
   Case vtIntegerUnsigned
      sResult = "vtIntegerUnsigned"
   Case vtLongIntegerUnsigned
      sResult = "vtLongIntegerUnsigned"
   Case vtQuadIntegerSigned
      sResult = "vtQuadIntegerSigned"
   Case vtQuadIntegerUnsigned
      sResult = "vtQuadIntegerUnsigned"
   Case vtLongInteger
      sResult = "vtLongInteger"
   Case vtDWord
      sResult = "vtDWord"
   Case vtVoid
      sResult = "vtVoid"
   Case vtHResult
      sResult = "vtHResult"
   Case vtPointer
      sResult = "vtPointer"
   Case vtSafeArray
      sResult = "vtSafeArray"
   Case vtCArray
      sResult = "vtCArray"
   Case vtUDT
      sResult = "vtUDT"
   Case vtLPStr
      sResult = "vtLPStr"
   Case vtLPWStr
      sResult = "vtLPWStr"
   Case vtFileTime
      sResult = "vtFileTime"
   Case vtBlob
      sResult = "vtBlob"
   Case vtStream
      sResult = "vtStream"
   Case vtStorage
      sResult = "vtStorage"
   Case vtStreamedObject
      sResult = "vtStreamedObject"
   Case vtStoredObject
      sResult = "vtStoredObject"
   Case vtBlobObject
      sResult = "vtBlobObject"
   Case vtClipboardFormat
      sResult = "vtClipboardFormat"
   Case vtCLSID
      sResult = "vtCLSID"
   Case vtVector
      sResult = "vtVector"
   Case vtArray
      sResult = "vtArray"
   Case vtByRef
      sResult = "vtByRef"
   Case (vtVector + 1) To (vtArray - 1)
      ' Vector of ...
      sResult = "vtVector of " & baTypeOfVariantToString(eValue - vtVector)
   Case (vtArray + 1) To (vtByRef - 1)
      ' Array of
      sResult = "vtArray of " & baTypeOfVariantToString(eValue - vtArray)
   Case Else
      sResult = "vtErrUnknown"
   End Select
   
   baTypeOfVariantToString = sResult
   
End Function
'==============================================================================

Private Function Trim0(ByVal sText As String) As String
   Trim0 = Trim$(Replace$(sText, Chr$(0), vbNullString))
End Function
'==============================================================================

