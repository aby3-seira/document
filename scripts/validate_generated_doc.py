import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
REPO = Path(__file__).resolve().parents[1]
DOC = REPO / 'output/SplunkDS_whitelist設定手順書.docx'


def main() -> None:
    if not DOC.exists():
        raise FileNotFoundError(f'生成物が見つかりません: {DOC}')

    with zipfile.ZipFile(DOC) as zf:
        xml = zf.read('word/document.xml')
    root = ET.fromstring(xml)
    text = '\n'.join((t.text or '') for t in root.findall('.//w:t', NS))

    required = [
        '3.1 Splunk Webへアクセス',
        '3.8 詳細タブで反映確認',
        '[serverClass:ishikari_uf]',
        'whitelist.0 = SplunkUF*',
        'whitelist.1 = DC001*',
    ]
    forbidden = ['UniversalForwarder', 'Windows', 'Linux', 'アンインストール作業']

    missing = [item for item in required if item not in text]
    remains = [item for item in forbidden if item in text]

    if missing or remains:
        if missing:
            print('Missing:', ', '.join(missing))
        if remains:
            print('Unexpected:', ', '.join(remains))
        raise SystemExit(1)

    print('Validation passed.')


if __name__ == '__main__':
    main()
