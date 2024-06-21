import os
import img2pdf
import sys

name = sys.argv[1]

dirname = "images"
imgs = []
for r, _, f in os.walk(dirname):
	for fname in f:
		if not fname.endswith(".jpg"):
			continue
		imgs.append(os.path.join(r, fname))
with open("images/"+ name +".pdf","wb") as f:
	f.write(img2pdf.convert(imgs))
