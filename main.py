from xml.dom.minidom import CharacterData


from define import getspreadsheet

characterlist = ['三条詩乃', '今神舞', '湊ラン', '相馬タクト', '品川天馬', '愛野葵']

for charactername in characterlist:
    print(f'{charactername}のチャットパレットを生成しています')
    getspreadsheet(charactername).exportchatpalette()
    print(f'{charactername}のチャットパレットの生成が完了しました')

print('全ての工程が終了しました')
