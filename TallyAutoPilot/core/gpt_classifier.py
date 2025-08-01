
import openai, pytesseract, pdf2image, tempfile
from PIL import Image

def extract_from_invoice(file, api_key):
    openai.api_key=api_key
    ext = os.path.splitext(file.name)[1].lower()
    # Save temp
    temp = tempfile.NamedTemporaryFile(delete=False,suffix=ext)
    temp.write(file.read()); temp.flush()
    pages = pdf2image.convert_from_path(temp.name) if ext==".pdf" else [Image.open(temp.name)]
    txt=""
    for img in pages:
        txt+=pytesseract.image_to_string(img)
    prompt=f"Extract PartyName, Amount, Date YYYY-MM-DD from invoice text:\n{txt}"
    res=openai.ChatCompletion.create(model="gpt-4",messages=[{"role":"user","content":prompt}])
    return eval(res.choices[0].message.content)
