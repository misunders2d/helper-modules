import os

excluded_collections = [
    'Cotton 300TC Percale Sheet Set','Egyptian Cotton Striped Bed Sheet Set',
    'Cotton Flannel Pillowcases','Decorative Pillow','All-Year-Round Throws',
    'Down Alternative Comforter','Faux Rabbit Fur Area Rug','Comforter',
    'Bed Skirt','3pc Microfiber Bed Sheet Set Full','Cotton Pillowcases',
    'Blackout Curtains','Cotton Flannel Fitted Sheet','Faux Fur Throw','Plush Coverlet Set',
    '3pc Microfiber Bed Sheet Set Queen','6 PC Egyptian Cotton Striped Bed Sheet Set',
    'Faux Cachemire Acrylic Throw Blanket','Cotton 300TC Percale Pillowcase',
    'Acrylic Knit Sherpa Throw Blanket','Cotton Quilt','Cotton 300TC Sateen Sheet Set',
    'Egyptian Cotton Bed Sheet Set','Jersey Cotton Quilt','Bundles','Pillow inserts'
    ]
user_folder = os.path.join(os.path.expanduser('~'), 'temp')
if not os.path.isdir(user_folder):
    os.makedirs(user_folder)

secrets_folder = os.path.join(os.getcwd(),'.secrets')
if not os.path.isdir(secrets_folder):
    os.makedirs(secrets_folder)