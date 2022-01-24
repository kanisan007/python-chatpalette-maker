import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
import time

class getspreadsheet:
    def __init__(self, sheetname:str) -> None:
        # sheetに触るための準備
        # gitで公開するために一部内容を差し替え
        self.scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('karioki.json', self.scope)
        self.gc = gspread.authorize(self.credentials)
        self.SPREADSHEET_KEY = '#############################################'
        self.worksheet = self.gc.open_by_key(self.SPREADSHEET_KEY).worksheet(sheetname)
        self.sheetname = sheetname

# 能力値の取得
    def getabilityscore(self) -> None:
        # 基礎ステータス
        print('基礎ステータスを取得中...')
        self.str = self.worksheet.acell('D6').value
        self.con = self.worksheet.acell('E6').value
        self.pow = self.worksheet.acell('F6').value
        self.dex = self.worksheet.acell('G6').value
        self.app = self.worksheet.acell('H6').value
        self.siz = self.worksheet.acell('I6').value
        self.int = self.worksheet.acell('J6').value
        self.edu = self.worksheet.acell('K6').value
        self.luck = self.worksheet.acell('R6').value
        time.sleep(9)
        print('基礎ステータスの取得完了')

    def getcombatskills(self) -> None:
        # 戦闘技能
        # 回避, 近接戦闘（基礎）, 遠隔戦闘（基礎）, 武道
        print('戦闘技能を取得中...')
        self.dodge = self.worksheet.acell('I16').value
        self.melee = self.worksheet.acell('I17').value
        self.throw = self.worksheet.acell('I19').value
        self.combatsarts = self.worksheet.acell('I18').value
        self.magic = self.worksheet.acell('I20').value
        time.sleep(5)
        print('戦闘技能の取得完了')

    def getsearchskills(self) -> None:
        # 探索技能
        # 応急手当, 鍵開け, 手さばき, 聞き耳, 隠密, 精神分析, 追跡, 登攀, 図書館, 目星, 鑑定, ヒプノーシス
        print('探索技能を取得中...')
        self.firstaid = self.worksheet.acell('S16').value
        self.locksmith = self.worksheet.acell('S17').value
        self.handling = self.worksheet.acell('S18').value
        self.listen = self.worksheet.acell('S19').value
        self.spy = self.worksheet.acell('S20').value
        self.psychoanalysis = self.worksheet.acell('S21').value
        self.Track = self.worksheet.acell('S22').value
        self.climb = self.worksheet.acell('S23').value
        self.libraryuse = self.worksheet.acell('S24').value
        self.spothidden = self.worksheet.acell('S25').value
        self.appraise = self.worksheet.acell('S26').value
        self.hypnosis = self.worksheet.acell('S27').value
        time.sleep(12)
        print('探索技能の取得完了')

    def getintellectskills(self) -> None:
        # 知識技能
        # 医学, オカルト, 芸術, 経理, 考古学, コンピューター, 一般教養学, 心理学, 人類学, 電子工学, 自然, 法律, 歴史, サバイバル, 科学
        print('知識技能を取得中...')
        self.medicine = self.worksheet.acell('I34').value
        self.occult = self.worksheet.acell('I35').value
        self.art = self.worksheet.acell('I36').value
        self.accounting = self.worksheet.acell('I37').value
        self.archaeology = self.worksheet.acell('I38').value
        self.computeruse = self.worksheet.acell('I39').value
        self.generaleducation = self.worksheet.acell('I40').value
        self.psychology = self.worksheet.acell('I41').value
        self.anthropology = self.worksheet.acell('I42').value
        self.electronics = self.worksheet.acell('I43').value
        self.nature = self.worksheet.acell('I44').value
        self.history = self.worksheet.acell('I45').value
        self.law = self.worksheet.acell('I46').value
        self.survival = self.worksheet.acell('I47').value
        self.science = self.worksheet.acell('I48').value
        time.sleep(15)
        print('知識技能の取得完了')

    def getbehavioralskills(self) -> None:
        # 行動技能
        # 運転, 機械修理, 重機械操作, 乗馬, 水泳, 制作, 操縦, 跳躍, 電気修理, ナビゲート, 変装
        print('行動技能を取得中...')
        self.driveautomobile = self.worksheet.acell('S34').value
        self.mechanicalrepair = self.worksheet.acell('S35').value
        self.operateheavymachine = self.worksheet.acell('S36').value
        self.ride = self.worksheet.acell('S37').value
        self.swim = self.worksheet.acell('S38').value
        self.craft = self.worksheet.acell('S39').value
        self.pilot = self.worksheet.acell('S40').value
        self.jump = self.worksheet.acell('S41').value
        self.electricalrepair = self.worksheet.acell('S42').value
        self.navigation = self.worksheet.acell('S43').value
        self.disguise = self.worksheet.acell('S44').value
        time.sleep(11)
        print('行動技能の取得完了')

    def getnegotiatingskills(self) -> None:
        # 交渉技能
        # 言いくるめ, 信用, 説得, 母国語, 威圧, 魅惑, 外国語
        print('交渉技能を取得中...')
        self.fasttalk = self.worksheet.acell('S49').value
        self.creditrating = self.worksheet.acell('S50').value
        self.persuade = self.worksheet.acell('S51').value
        self.ownlanguage = self.worksheet.acell('S52').value
        self.coercion = self.worksheet.acell('S53').value
        self.charm = self.worksheet.acell('S54').value
        self.otherlanguage = self.worksheet.acell('S55').value
        time.sleep(7)
        print('交渉技能の取得完了')

    def getadditionalskills(self) -> None:
        self.additionalskill1name = self.worksheet.acell('A62').value
        self.additionalskill1 = self.worksheet.acell('I62').value

        if self.additionalskill1name is None:
            self.additionalskill1name = ''
            self.additionalskill1 = ''

        self.additionalskill2name = self.worksheet.acell('A63').value
        self.additionalskill2 = self.worksheet.acell('I63').value

        if self.additionalskill2name is None:
            self.additionalskill2name = ''
            self.additionalskill2 = ''

        self.additionalskill3name = self.worksheet.acell('A64').value
        self.additionalskill3 = self.worksheet.acell('I64').value

        if self.additionalskill3name is None:
            self.additionalskill3name = ''
            self.additionalskill3 = ''

        self.additionalskill4name = self.worksheet.acell('A65').value
        self.additionalskill4 = self.worksheet.acell('I65').value

        if self.additionalskill4name is None:
            self.additionalskill4name = ''
            self.additionalskill4 = ''


    def getallskills(self) -> None:
        # 全部まとめて呼ぶときはこっち
        self.getabilityscore()
        self.getcombatskills()
        self.getsearchskills()
        self.getintellectskills()
        self.getbehavioralskills()
        self.getnegotiatingskills()
        self.getadditionalskills()
        print('すべての技能値を取得しました')

    def createchatpalette(self) -> str:
        # チャットパレットを作る
        self.getallskills()
        text = '''
CC<={} 【回避】
CC<={} 【近接戦闘(基礎)】
CC<={} 【遠隔戦闘(基礎)】

CC<={} 【アイデア】
CC<={} 【知識】
CC<={} 【幸運】

CC<={} 【応急手当】
CC<={} 【鍵開け】
CC<={} 【手さばき】
CC<={} 【聞き耳】
CC<={} 【隠密】
CC<={} 【精神分析】
CC<={} 【追跡】
CC<={} 【登攀】
CC<={} 【図書館】
CC<={} 【目星】
CC<={} 【鑑定】
CC<={} 【ヒプノーシス】

CC<={} 【医学】
CC<={} 【オカルト】
CC<={} 【芸術】
CC<={} 【経理】
CC<={} 【考古学】
CC<={} 【コンピューター】
CC<={} 【一般教養学】
CC<={} 【心理学】
CC<={} 【人類学】
CC<={} 【電子工学】
CC<={} 【自然】
CC<={} 【法律】
CC<={} 【歴史】
CC<={} 【サバイバル】
CC<={} 【科学】

CC<={} 【{}】
CC<={} 【{}】
CC<={} 【{}】
CC<={} 【{}】

CC<={} 【運転】
CC<={} 【機械修理】
CC<={} 【重機械操作】
CC<={} 【乗馬】
CC<={} 【水泳】
CC<={} 【製作】
CC<={} 【操縦】
CC<={} 【跳躍】
CC<={} 【電気修理】
CC<={} 【ナビゲート】
CC<={} 【変装】

CC<={} 【言いくるめ】
CC<={} 【信用】
CC<={} 【説得】
CC<={} 【母国語】
CC<={} 【威圧】
CC<={} 【魅惑】
CC<={} 【言語】

CC<={} 【STRロール】
CC<={} 【CONロール】
CC<={} 【POWロール】
CC<={} 【DEXロール】
CC<={} 【APPロール】
CC<={} 【SIZロール】
        '''
        skillsdata = text.format(
            self.dodge,\
            self.melee,\
            self.throw,\
            \
            self.int,\
            self.edu,\
            self.luck,\
            \
            self.firstaid,\
            self.locksmith,\
            self.handling,\
            self.listen,\
            self.spy,\
            self.psychoanalysis,\
            self.Track,\
            self.climb,\
            self.libraryuse,\
            self.spothidden,\
            self.appraise,\
            self.hypnosis,\
            \
            self.medicine,\
            self.occult,\
            self.art,\
            self.accounting,\
            self.archaeology,\
            self.computeruse,\
            self.generaleducation,\
            self.psychology,\
            self.anthropology,\
            self.electronics,\
            self.nature,\
            self.history,\
            self.law,\
            self.survival,\
            self.science,\
            \
            self.additionalskill1,\
            self.additionalskill1name,\
            self.additionalskill2,\
            self.additionalskill2name,\
            self.additionalskill3,\
            self.additionalskill3name,\
            self.additionalskill4,\
            self.additionalskill4name,\
            \
            self.driveautomobile,\
            self.mechanicalrepair,\
            self.operateheavymachine,\
            self.ride,\
            self.swim,\
            self.craft,\
            self.pilot,\
            self.jump,\
            self.electricalrepair,\
            self.navigation,\
            self.disguise,\
            \
            self.fasttalk,\
            self.creditrating,\
            self.persuade,\
            self.ownlanguage,\
            self.coercion,\
            self.charm,\
            self.otherlanguage,\
            \
            self.str,\
            self.con,\
            self.pow,\
            self.dex,\
            self.app,\
            self.siz\
            )
        # .formatではじかれるので後から中括弧を使う分を追加する
        chatpalette = '''
1d{レベル}+3+{db}+{威力}
CC<={SAN} 【SAN値チェック】
        ''' + skillsdata

        return chatpalette

    def exportchatpalette(self) -> None:
        # .txtに出力する
        palettedata = self.createchatpalette()
        f = open(self.sheetname + '.txt', 'w', encoding='UTF-8')
        f.write(palettedata)
        f.close()


