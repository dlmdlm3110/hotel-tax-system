import os
from PIL import Image

d = os.path.dirname(os.path.abspath(__file__))
img = Image.open(os.path.join(d, '宿泊税スマートロゴ.png')).convert('RGBA')
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
frames = [img.resize(s, Image.LANCZOS) for s in sizes]
frames[0].save(os.path.join(d, 'appicon.ico'),
               format='ICO', sizes=sizes, append_images=frames[1:])
print('appicon.ico 作成完了')
