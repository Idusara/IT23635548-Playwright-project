import openpyxl
import os

data = [
    # Question forms
    ("Neg_0001", "oya ada class enawada nadda?", "ඔයා අද පන්තියට එනවද නැද්ද?", "Question forms", "Rationale: The overall message is a question"),
    ("Neg_0002", "me dawas wala mokada karanne?", "මේ දවස් වල මොකද කරන්නෙ?", "Question forms", "Rationale: The overall message is a question"),
    
    # Command forms
    ("Neg_0003", "ikmanata me wede iwara karanna.", "ඉක්මනට මේ වැඩේ ඉවර කරන්න.", "Command forms", "Rationale: The input is a direct command"),
    ("Neg_0004", "kageda kiyala balala mata kiyanna.", "කාගෙද කියලා බලලා මට කියන්න.", "Command forms", "Rationale: The input commands someone to do something"),
    
    # Greetings
    ("Neg_0005", "subha udasanak wewa hamotama!", "සුභ උදෑසනක් වේවා හැමෝටම!", "Greetings", "Rationale: Contains a common morning greeting"),
    ("Neg_0006", "obata subha rathriyak!", "ඔබට සුභ රාත්‍රියක්!", "Greetings", "Rationale: Contains a night time greeting"),
    
    # Requests
    ("Neg_0007", "karunakara mage message ekata reply karanna.", "කරුණාකර මගේ මැසේජ් එකට රිප්ලයි කරන්න.", "Requests", "Evidence: 'karunakara mage message ekata reply karanna'"),
    ("Neg_0008", "mata podi udawwak karanna puluwanda?", "මට පොඩි උදව්වක් කරන්න පුළුවන්ද?", "Requests", "Evidence: 'mata podi udawwak karanna puluwanda?'"),
    
    # Responses
    ("Neg_0009", "ow, mama eka heta karannam.", "ඔව්, මම ඒක හෙට කරන්නම්.", "Responses", "Evidence: 'ow' indicates an affirmative response"),
    ("Neg_0010", "naa, mata eka karanna beha.", "නෑ, මට ඒක කරන්න බැහැ.", "Responses", "Evidence: 'naa' indicates a negative response"),
    
    # Repeated Words
    ("Neg_0011", "hemin hemin parissamen yanna.", "හෙමින් හෙමින් පරිස්සමෙන් යන්න.", "Repeated Words", "Evidence: The word 'hemin' is repeated"),
    ("Neg_0012", "hoda hoda dewal issarahata thiyanawa.", "හොඳ හොඳ දේවල් ඉස්සරහට තියෙනවා.", "Repeated Words", "Evidence: The word 'hoda' is repeated"),
    
    # Inputs with Punctuation Marks
    ("Neg_0013", "mokada wenne...? man balan innawa!", "මොකද වෙන්නේ...? මං බලන් ඉන්නවා!", "Inputs with Punctuation Marks", "Rationale: Contains punctuation such as ...? and !"),
    ("Neg_0014", "oh! eka godak lassanai.", "ඕහ්! ඒක ගොඩක් ලස්සනයි.", "Inputs with Punctuation Marks", "Rationale: Contains punctuation such as ! and ."),
    
    # Romanization / Spelling Variants
    ("Neg_0015", "mn oyaata psse katha krnnm.", "මං ඔයාට පස්සේ කතා කරන්නම්.", "Romanization / Spelling Variants", "Rationale: The words are written with missing vowels like 'mn', 'psse', 'krnnm'"),
    ("Neg_0016", "kohmada oyaata dan saniipada?", "කොහොමද ඔයාට දැන් සනීපද?", "Romanization / Spelling Variants", "Rationale: The word 'kohmada' is a spelling variant"),
    
    # Isolated English Word Insertions in Singlish
    ("Neg_0017", "ada mama game ekak gahuwa.", "අද මම ගේම් එකක් ගැහුවා.", "Isolated English Word Insertions in Singlish", "Rationale: Single word 'game' is inserted"),
    ("Neg_0018", "eyage style eka harima wenas.", "එයාගේ ස්ටයිල් එක හරිම වෙනස්.", "Isolated English Word Insertions in Singlish", "Rationale: Single word 'style' is inserted"),
    
    # Multi-Word English Phrases in Singlish
    ("Neg_0019", "mama hitan hitiye this is a joke kiyala.", "මම හිතන් හිටියේ මේක විහිළුවක් කියලා.", "Multi-Word English Phrases in Singlish", "Rationale: Phrase 'this is a joke' is inserted"),
    ("Neg_0020", "eka nam out of this world experience ekak.", "ඒක නම් out of this world එක්ස්පීරියන්ස් එකක්.", "Multi-Word English Phrases in Singlish", "Rationale: Phrase 'out of this world' is inserted"),
    
    # English Digital Terms in Singlish
    ("Neg_0021", "mage router eka wada na.", "මගේ රවුටර් එක වැඩ නෑ.", "English Digital Terms in Singlish", "Rationale: Input includes the digital term 'router'"),
    ("Neg_0022", "oyage account eka log out wela.", "ඔයාගේ එකවුන්ට් එක ලොග් අවුට් වෙලා.", "English Digital Terms in Singlish", "Rationale: Input includes digital terms 'account' and 'log out'"),
    
    # Platform/App Names in Singlish
    ("Neg_0023", "mama eka Facebook eke dakka.", "මම ඒක ෆේස්බුක් එකේ දැක්කා.", "Platform/App Names in Singlish", "Rationale: Input includes platform name 'Facebook'"),
    ("Neg_0024", "Netflix balanna thibba nam maru.", "නෙට්ෆ්ලික්ස් බලන්න තිබ්බා නම් මරු.", "Platform/App Names in Singlish", "Rationale: Input includes platform name 'Netflix'"),
    
    # English Abbreviations/Acronyms in Singlish
    ("Neg_0025", "oyage CV eka ewanna puluwanda?", "ඔයාගේ සීවී එක එවන්න පුළුවන්ද?", "English Abbreviations/Acronyms in Singlish", "Rationale: Input includes abbreviation 'CV'"),
    ("Neg_0026", "CEO kiwwa meeting eka cancel kiyala.", "සීඊඕ කිව්වා මීටින් එක කැන්සල් කියලා.", "English Abbreviations/Acronyms in Singlish", "Rationale: Input includes acronym 'CEO'"),
    
    # English Clipped Forms in Singlish
    ("Neg_0027", "mage lab eke weda tikak thiyenawa.", "මගේ ලැබ් එකේ වැඩ ටිකක් තියෙනවා.", "English Clipped Forms in Singlish", "Rationale: Input includes clipped form 'lab'"),
    ("Neg_0028", "api ad ad eka damu.", "අපි ඇඩ් එක දාමු.", "English Clipped Forms in Singlish", "Rationale: Input includes clipped form 'ad'"),
    
    # Place Names Embedded in Singlish
    ("Neg_0029", "api heta Kandy yanawa.", "අපි හෙට මහනුවර යනවා.", "Place Names Embedded in Singlish", "Rationale: Input includes place name 'Kandy'"),
    ("Neg_0030", "Colombo wala traffic godak wadi.", "කොළඹ වල ට්‍රැෆික් ගොඩක් වැඩියි.", "Place Names Embedded in Singlish", "Rationale: Input includes place name 'Colombo'"),
    
    # Person Names Embedded in Singlish
    ("Neg_0031", "Pathum ayya ada office awe na.", "පැතුම් අයියා අද ඔෆිස් ආවේ නෑ.", "Person Names Embedded in Singlish", "Rationale: Input includes person name 'Pathum'"),
    ("Neg_0032", "Kamali ge wedding eka langai.", "කමලිගේ වෙඩින් එක ළඟයි.", "Person Names Embedded in Singlish", "Rationale: Input includes person name 'Kamali'"),
    
    # Inputs with Numbers and Numeric Suffixes
    ("Neg_0033", "mata 100k withara denna puluwanda?", "මට සීයක් විතර දෙන්න පුළුවන්ද?", "Inputs with Numbers and Numeric Suffixes", "Rationale: Input includes number with suffix '100k'"),
    ("Neg_0034", "eka 2nd time ekata wune.", "ඒක දෙවැනි පාරට වුණේ.", "Inputs with Numbers and Numeric Suffixes", "Rationale: Input includes number with suffix '2nd'"),
    
    # Inputs with Currency
    ("Neg_0035", "meke gana $50k withara wenawa.", "මේකේ ගාණ ඩොලර් 50ක් විතර වෙනවා.", "Inputs with Currency", "Rationale: Input includes currency symbol '$50'"),
    ("Neg_0036", "eka LKR 5000kata denna baha.", "ඒක රුපියල් 5000කට දෙන්න බැහැ.", "Inputs with Currency", "Rationale: Input includes currency code 'LKR 5000'"),
    
    # Inputs with Time Formats
    ("Neg_0037", "api 8.30am walata set wemu.", "අපි උදේ 8.30ට සෙට් වෙමු.", "Inputs with Time Formats", "Rationale: Input includes time format '8.30am'"),
    ("Neg_0038", "meeting eka 4:15 PM ta iwara wenawa.", "මීටින් එක හවස 4:15ට ඉවර වෙනවා.", "Inputs with Time Formats", "Rationale: Input includes time format '4:15 PM'"),
    
    # Inputs with Dates
    ("Neg_0039", "mage upandinaya 2001-05-20.", "මගේ උපන්දිනේ 2001-05-20.", "Inputs with Dates", "Rationale: Input includes date '2001-05-20'"),
    ("Neg_0040", "January 1st weni da api meet wemu.", "ජනවාරි 1 වැනිදා අපි මීට් වෙමු.", "Inputs with Dates", "Rationale: Input includes date 'January 1st'"),
    
    # Inputs with Unit of Measurements
    ("Neg_0041", "mage usa 170cm withara.", "මගේ උස සෙන්ටිමීටර් 170ක් විතර.", "Inputs with Unit of Measurements", "Rationale: Input includes measurement unit '170cm'"),
    ("Neg_0042", "bat eke bara 2kg k withara thiyenawa.", "බැට් එකේ බර කිලෝග්‍රෑම් 2ක් විතර තියෙනවා.", "Inputs with Unit of Measurements", "Rationale: Input includes measurement unit '2kg'"),
    
    # Inputs with Slang and Casual Phrasing
    ("Neg_0043", "patta machan, wede niyameta giya.", "පට්ට මචං, වැඩේ නියමෙට ගියා.", "Inputs with Slang and Casual Phrasing", "Rationale: Input includes slang words 'patta' and 'machan'"),
    ("Neg_0044", "ammo eka nam supiri bhandayak.", "අම්මෝ ඒක නම් සුපිරි භාණ්ඩයක්.", "Inputs with Slang and Casual Phrasing", "Rationale: Input includes casual phrasing 'ammo' and 'supiri'"),
    
    # Online Indentifiers in Singlish
    ("Neg_0045", "@nimal ayye, poddak me link eka balanna youtube.com.", "@නිමල් අයියේ, පොඩ්ඩක් මේ ලින්ක් එක බලන්න youtube.com.", "Online Indentifiers in Singlish", "Rationale: Input includes identifiers '@nimal' and 'youtube.com'"),
    ("Neg_0046", "mage email eka test@email.com walata ewanna.", "මගේ ඊමේල් එක test@email.com වලට එවන්න.", "Online Indentifiers in Singlish", "Rationale: Input includes email address 'test@email.com'"),
    
    # Inputs Containing Emojis
    ("Neg_0047", "ammo mata hina gihin marenawa 😂", "අම්මෝ මට හිනා ගිහින් මැරෙනවා 😂", "Inputs Containing Emojis", "Rationale: Input contains the laughing emoji 😂"),
    ("Neg_0048", "eka nam hena duk seen ekak 😢", "ඒක නම් හේන දුක් සීන් එකක් 😢", "Inputs Containing Emojis", "Rationale: Input contains the crying emoji 😢"),
    
    # Extra 2 cases
    ("Neg_0049", "eya harima busy this week nisa call kare na.", "එයා හරිම කාර්යබහුලයි මේ සතියේ නිසා කෝල් කරේ නෑ.", "Multi-Word English Phrases in Singlish", "Rationale: Phrase 'this week' is inserted"),
    ("Neg_0050", "kiyala kiyala epa wela thiyenne.", "කියලා කියලා එපා වෙලා තියෙන්නේ.", "Repeated Words", "Evidence: The word 'kiyala' is repeated")
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = " Test cases"

headers = ["TC ID", "Input length type", "Input", "Expected output", "Actual output", "Status", "Cat", "Rat"]
ws.append(headers)

for item in data:
    tc_id, singlish, sinhala, cat, rat = item
    length = len(singlish)
    if length <= 30:
        l_type = "S"
    elif length <= 299:
        l_type = "M"
    else:
        l_type = "L"
    
    row = [tc_id, l_type, singlish, sinhala, "", "", cat, rat]
    ws.append(row)

wb.save("temp.xlsx")
print("Saved 50 test cases to temp.xlsx")
