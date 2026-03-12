import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{%s}' % NS['w']
ET.register_namespace('w', NS['w'])

REPO = Path(__file__).resolve().parents[1]
TEMPLATE = REPO / 'source/template/UniversalForwarderインストールマニュアル.docx'
OUTPUT = REPO / 'output/SplunkDS_whitelist設定手順書.docx'


def paragraph_text(p):
    return ''.join((t.text or '') for t in p.findall('.//w:t', NS)).strip()


def make_paragraph(text, style=None):
    p = ET.Element(W + 'p')
    ppr = ET.SubElement(p, W + 'pPr')
    if style:
        ps = ET.SubElement(ppr, W + 'pStyle')
        ps.set(W + 'val', style)
    r = ET.SubElement(p, W + 'r')
    t = ET.SubElement(r, W + 't')
    if text.startswith(' ') or text.endswith(' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return p


def make_blank():
    return ET.Element(W + 'p')


def load_template(template_path: Path) -> dict:
    if not template_path.exists():
        raise FileNotFoundError(f'テンプレートが見つかりません: {template_path}')
    if template_path.stat().st_size == 0:
        raise RuntimeError(f'テンプレートが空ファイルです: {template_path}')

    with zipfile.ZipFile(template_path, 'r') as zin:
        if 'word/document.xml' not in zin.namelist():
            raise RuntimeError('テンプレートが不正です: word/document.xml が存在しません。')
        return {name: zin.read(name) for name in zin.namelist()}


def main() -> None:
    files = load_template(TEMPLATE)

    root = ET.fromstring(files['word/document.xml'])
    body = root.find('w:body', NS)
    children = list(body)

    # Replace title strings on cover and other remaining runs.
    for t in root.findall('.//w:t', NS):
        if t.text:
            t.text = t.text.replace('UniversalForwarderインストールマニュアル', 'Splunk Deployment Server whitelist設定手順書')
            t.text = t.text.replace('インストールマニュアル', 'whitelist設定手順書')
            t.text = t.text.replace('UniversalForwarder', 'Splunk Deployment Server')

    # Clear existing static TOC text to avoid stale entries; field can be updated in Word.
    for sdt in root.findall('.//w:sdt', NS):
        for t in sdt.findall('.//w:t', NS):
            if t.text:
                t.text = ''

    # Remove existing operation chapters from first "Linux" heading onward.
    start_idx = None
    for i, ch in enumerate(children):
        if ch.tag == W + 'p' and paragraph_text(ch) == 'Linux':
            start_idx = i
            break

    if start_idx is None:
        raise RuntimeError('テンプレート内で本文開始位置（Linux見出し）が見つかりませんでした。')

    for ch in children[start_idx:-1]:
        body.remove(ch)

    new_content = [
        make_paragraph('1. 概要', style='a9'),
        make_paragraph('本手順書は、Splunk Deployment Server の ServerClass whitelist を Splunk Web から設定するための手順を説明するものです。対象は SplunkDS-pr です。'),
        make_blank(),
        make_paragraph('2. 事前確認事項', style='a9'),
        make_paragraph('以下を事前に確認してください。'),
        make_paragraph('1) Splunk Web にログイン可能であること'),
        make_paragraph('2) 対象 ServerClass（ishikari_uf）が存在すること'),
        make_paragraph('3) 追加対象のクライアント名／ホスト名／IP アドレス／DNS 名が分かっていること'),
        make_blank(),
        make_paragraph('3. whitelist設定手順', style='a9'),
        make_paragraph('3.1 Splunk Webへアクセス', style='1'),
        make_paragraph('1) Web ブラウザから Splunk Web へアクセスし、管理者アカウントでログインします。'),
        make_blank(),
        make_paragraph('3.2 エージェント管理を開く', style='1'),
        make_paragraph('1) 画面上部の「設定」をクリックします。'),
        make_paragraph('2) 「エージェント管理」をクリックします。'),
        make_paragraph('【画像プレースホルダ：エージェント管理画面】'),
        make_blank(),
        make_paragraph('3.3 サーバークラスを開く', style='1'),
        make_paragraph('1) エージェント管理画面の上部タブから「サーバークラス」をクリックします。'),
        make_paragraph('【画像プレースホルダ：サーバークラス画面】'),
        make_blank(),
        make_paragraph('3.4 ServerClass(ishikari_uf)を選択', style='1'),
        make_paragraph('1) 一覧から ServerClass「ishikari_uf」を選択します。'),
        make_blank(),
        make_paragraph('3.5 フォワーダーの編集', style='1'),
        make_paragraph('1) 対象 ServerClass 画面で「フォワーダーの編集」をクリックします。'),
        make_paragraph('【画像プレースホルダ：フォワーダー編集画面】'),
        make_blank(),
        make_paragraph('3.6 whitelist追加', style='1'),
        make_paragraph('1) 「含める（必須）」欄に対象をカンマ区切りで入力します。'),
        make_paragraph('2) 入力例：SplunkUF*,DC001*'),
        make_paragraph('3) 「指定のホスト名の一致」にチェックが入っていることを確認します。'),
        make_blank(),
        make_paragraph('3.7 保存', style='1'),
        make_paragraph('1) 「保存」をクリックします。'),
        make_blank(),
        make_paragraph('3.8 詳細タブで反映確認', style='1'),
        make_paragraph('1) 保存後、「詳細」タブを開きます。'),
        make_paragraph('2) 対象スタンザ [serverClass:ishikari_uf] に whitelist が追加されていることを確認します。'),
        make_paragraph('【画像プレースホルダ：詳細タブ確認画面】'),
        make_blank(),
        make_paragraph('4. 確認ポイント', style='a9'),
        make_paragraph('保存後、詳細タブの [serverClass:ishikari_uf] に whitelist が反映されていることを確認してください。'),
        make_blank(),
        make_paragraph('5. 参考', style='a9'),
        make_paragraph('serverclass.conf の設定イメージ：'),
        make_paragraph('[serverClass:ishikari_uf]'),
        make_paragraph('whitelist.0 = SplunkUF*'),
        make_paragraph('whitelist.1 = DC001*'),
    ]

    insert_at = len(list(body)) - 1
    for p in new_content:
        body.insert(insert_at, p)
        insert_at += 1

    files['word/document.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(f'Generated: {OUTPUT}')


if __name__ == '__main__':
    main()
