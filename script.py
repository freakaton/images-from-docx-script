import os
import zipfile

root_dir = os.getcwd()
docs = os.listdir()
#sorting list by only .zip, that earlier were .docx
for doc in docs:
    if not doc.endswith('.docx'):
        docs.remove(doc)

#extracting images
for doc in docs:
    try:
        zipf = zipfile.ZipFile(doc, 'r')
        for file in zipf.namelist():
            if file.endswith('.png') or file.endswith('.jpg'):
                image = zipf.getinfo(file)
                zipf.extract(image, 'tmp/' + zipf.filename.replace('.docx', '').strip())
    except zipfile.BadZipfile:
        print(doc, ' not a zip-file, passing..')
os.chdir(os.path.join(os.getcwd(), 'tmp'))

#rename images and change their directory
for name in os.listdir():
    path_to_image = os.path.join(name, 'word', 'media')
    name_of_image = os.listdir(os.path.join(os.getcwd(), path_to_image))[0]
    path_to_image = os.path.join(path_to_image, name_of_image)
    os.renames(path_to_image, os.path.join(root_dir, 'result', name+'.png'))


os.chdir(os.path.join(root_dir, 'result'))
for im in os.listdir():
    os.rename(im, im.replace(' .png', '.png'))
os.rmdir(os.path.join(root_dir, 'tmp'))
    
    

