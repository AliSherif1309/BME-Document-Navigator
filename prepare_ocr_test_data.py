"""Download a checksum-pinned official English model for CI/tests, not shipping."""
import hashlib
from pathlib import Path
from urllib.request import urlopen

REVISION = '87416418657359cb625c412a48b6e1d6d41c29bd'
SHA256 = '7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2'


def main():
    folder = Path(__file__).resolve().parent / '.ocr-testdata'
    folder.mkdir(exist_ok=True)
    target = folder / 'eng.traineddata'
    if target.is_file() and hashlib.sha256(target.read_bytes()).hexdigest()==SHA256:
        print('Verified English OCR test data.'); return
    url = f'https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/{REVISION}/eng.traineddata'
    with urlopen(url,timeout=30) as response: content = response.read(16000000)
    if hashlib.sha256(content).hexdigest()!=SHA256:
        raise RuntimeError('Official OCR test data did not match the pinned checksum.')
    target.write_bytes(content)
    print('Downloaded verified English OCR test data. Do not include it in application artifacts.')


if __name__=='__main__': main()
