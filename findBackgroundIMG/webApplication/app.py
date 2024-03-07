from flask import Flask, render_template, jsonify
import csv

app = Flask(__name__)

def load_images_from_csv(csv_file_path):
    images = {}
    with open(csv_file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            img_hash = row['Image Hash']
            if img_hash not in images:
                images[img_hash] = {'img_path': row['Image File'], 'pptx_paths': [row['PPTX File']]}
            else:
                if row['PPTX File'] not in images[img_hash]['pptx_paths']:
                    images[img_hash]['pptx_paths'].append(row['PPTX File'])
    return list(images.values())

@app.route('/')
def index():
    images_data = load_images_from_csv('source/image_ppt_mapping.csv')
    return render_template('index.html', images_data=images_data)

if __name__ == '__main__':
    app.run(debug=True)
