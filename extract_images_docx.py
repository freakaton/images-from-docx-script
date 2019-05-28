import os
import shutil
import zipfile


ROOT_DIR = os.getcwd()
IMAGE_FORMATS = ['png', 'jpg', 'jpeg', 'gif', 'bmp']

try:
    os.mkdir('tmp/')
except OSError:
    pass
try:
    os.mkdir('result/')
except OSError:
    pass

# Sorting list by only .zip, that earlier were .docx
docs = filter(lambda a: a.endswith('.docx'), os.listdir('.'))

# extracting images
for doc in docs:
    try:
        zip_file = zipfile.ZipFile(doc, 'r')
    except (zipfile.BadZipfile, IsADirectoryError):
        print(doc, ' not a zip-file, passing..')
        continue

    for file in zip_file.namelist():
        if file.split('.')[-1] in IMAGE_FORMATS:
            image = zip_file.getinfo(file)
            zip_file.extract(image, 'tmp/' + zip_file.filename.replace('.docx', '').strip())
os.chdir(os.path.join(os.getcwd(), 'tmp'))

# rename images and change their directory
for name in os.listdir('.'):
    path_to_image = os.path.join(name, 'word', 'media')
    name_of_image = os.listdir(os.path.join(os.getcwd(), path_to_image))[0]
    path_to_image = os.path.join(path_to_image, name_of_image)
    os.renames(path_to_image, os.path.join(ROOT_DIR, 'result', name+'.png'))

shutil.rmtree(os.path.join(ROOT_DIR, 'tmp'))
