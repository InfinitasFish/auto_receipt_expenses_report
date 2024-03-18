from pathlib import Path
import numpy as np
import torch
from PIL import Image, ImageDraw
from transformers import LayoutLMv2ForTokenClassification
from transformers import LayoutXLMProcessor
from yolo_pipeline import yolo_pipeline
import cv2
from pdf2image import convert_from_path
from additional_utils import fill_exel_report
import easyocr
import math


label_list = ['O', 'B-COMPANY', 'I-COMPANY', 'B-DATE', 'I-DATE', 'B-ADDRESS', 'I-ADDRESS', 'B-TOTAL', 'I-TOTAL']
id2label = {v: k for v, k in enumerate(label_list)}
label2id = {k: v for v, k in enumerate(label_list)}
label2color = {'O': 'gray', 'B-COMPANY': 'red', 'I-COMPANY': 'red', 'B-ADDRESS': 'orange',
               'I-ADDRESS': 'orange', 'B-DATE': 'violet', 'I-DATE': 'violet', 'B-TOTAL': 'green', 'I-TOTAL': 'green'}


def whole_pipeline(layout_model, processor, ocr_engine, yolo_model, filename):
    # checking for multiple receipts on image
    start_np_image = convert_image_to_numpy(filename)
    separated_images = yolo_pipeline(yolo_model, start_np_image)

    # applying pipeline for each image
    results = [] # dictionaries with all data
    extracted_data_results = [] # separate list for filling exel report
    for np_image in separated_images:
        predictions, predicted_boxes, extracted_data = layout_forward_pass(layout_model, processor, ocr_engine, np_image)
        if predictions is not None:
            extracted_data_results.append(extracted_data)
            annotated_image = get_annotated_image(np_image, predictions, predicted_boxes)
            qr_code_data = qr_code_data_extraction(np_image)
            combined_extracted_data = combine_informative_tokens(extracted_data)

            results.append({'combined_extracted_data': combined_extracted_data,
                            'annotated_image': annotated_image,
                            'qr_code_data': qr_code_data})

    fill_exel_report(extracted_data_results)

    return results


def initialize_layout_model():
    model = LayoutLMv2ForTokenClassification.from_pretrained("./layoutXLM_checkpoint",
                                                             id2label=id2label,
                                                             label2id=label2id,
                                                             num_labels=len(label_list))

    return model


def initialize_layout_processor():
    processor = LayoutXLMProcessor.from_pretrained("layoutXLM_checkpoint", apply_ocr=False)

    return processor


def initialize_easyocr_detector():
    return easyocr.Reader(['ru'])


def layout_forward_pass(layout_model, processor, ocr_engine, np_image):
    device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')

    layout_model.to(device)

    width, height = np_image.shape[1], np_image.shape[0]

    # using easyocr instead of tesseract
    ocr_results = ocr_engine.readtext(np_image, width_ths=0.1, y_ths=0.1)

    formatted_results = []

    if len(ocr_results) > 0:
        for result in ocr_results:
            points = result[0]

            p0, p2 = points[0], points[2]
            x0, y0, x2, y2 = max(0, p0[0]), max(0, p0[1]), min(p2[0], width), min(p2[1], height)

            norm_box = normalize_box_1000([x0, y0, x2, y2], height, width)
            formatted_results.append((norm_box, result[1]))

        words = [res[1] for res in formatted_results]
        boxes = [res[0] for res in formatted_results]

        encoding = processor(np_image, text=words, boxes=boxes, truncation=True,
                             return_offsets_mapping=True, return_tensors="pt",
                             padding="max_length", max_length=512)

        offset_mapping = encoding.pop('offset_mapping')

        for k, v in encoding.items():
            encoding[k] = v.to(device)

        outputs = layout_model(**encoding)
        predictions = outputs.logits.argmax(-1).squeeze().tolist()
        token_boxes = encoding.bbox.squeeze().tolist()

        is_subword = np.array(offset_mapping.squeeze().tolist())[:, 0] != 0

        true_predictions = [id2label[pred] for idx, pred in enumerate(predictions) if not is_subword[idx]]
        true_boxes = [unnormalize_1000_box(box, width, height) for idx, box in enumerate(token_boxes) if not is_subword[idx]]

        extracted_data = get_informative_tokens_text(encoding, offset_mapping, outputs, id2label, processor)

        return true_predictions, true_boxes, extracted_data
    else:
        return None, None, None


def get_annotated_image(np_image, predictions, boxes):

    pil_image = Image.fromarray(np_image)

    width, height = pil_image.size
    box_width = math.ceil(width * height / 10e6)
    font_size = box_width * 10

    image_draw = ImageDraw.Draw(pil_image)

    for prediction, box in zip(predictions, boxes):
        predicted_label = iob_to_label(prediction)
        if predicted_label != id2label[0]:
            image_draw.rectangle(box, outline=label2color[predicted_label], width=box_width)
            image_draw.text((box[0] + 10, box[1] - 20), text=predicted_label, fill=label2color[predicted_label], font_size=font_size)
        else:
            image_draw.rectangle(box, outline=label2color[predicted_label], width=box_width)

    return pil_image


def normalize_box_1000(box, height, width):

  x0, y0, x2, y2 = [int(p) for p in box]

  x0 = int(1000 * (x0 / width))
  x2 = int(1000 * (x2 / width))
  y0 = int(1000 * (y0 / height))
  y2 = int(1000 * (y2 / height))

  return [x0, y0, x2, y2]


def unnormalize_1000_box(bbox, width, height):
    return [
        int(width * (bbox[0] / 1000)),
        int(height * (bbox[1] / 1000)),
        int(width * (bbox[2] / 1000)),
        int(height * (bbox[3] / 1000)),
    ]


def get_informative_tokens_text(encoding, offset_mapping, outputs, id2label, processor):

    input_ids = encoding['input_ids'].squeeze().tolist()

    is_subword = np.array(offset_mapping.squeeze().tolist())[:,0] != 0

    predictions = outputs.logits.argmax(-1).squeeze().tolist()
    true_predictions = [id2label[pred] for idx, pred in enumerate(predictions) if not is_subword[idx]]

    full_words = []
    word_idx = []
    for i in range(len(is_subword) - 1):
        word_idx.append(i)
        if is_subword[i + 1]:
          continue
        else:
            full_words.append(''.join([processor.tokenizer.decode(input_ids[id]) for id in word_idx]))
            word_idx = []
    # adding 'end' token
    full_words.append('</s>')

    inf_pred_word_tuples = []
    for i in range(len(true_predictions)):
        if true_predictions[i] != id2label[0]:
            inf_pred_word_tuples.append((true_predictions[i], full_words[i]))

    return inf_pred_word_tuples


def combine_informative_tokens(extracted_data):
    # extracted_data is list of ( ()-TOKEN, WORD ) tuples

    data_combined_dict = {'COMPANY': [], 'ADDRESS': [], 'DATE': [], 'TOTAL': []}
    for (token, word) in extracted_data:
        # skip <s> </s> tokens
        if word not in ('<s>', '</s>'):
            part_token = token.split('-')[1]
            data_combined_dict[part_token].append(word)

    # ;)
    # data_combined_dict['DATE'] = '.'.join(data_combined_dict['DATE'])

    return data_combined_dict


def iob_to_label(label):
    if not label:
      return 'O'
    return label


def qr_code_data_extraction(np_image):
    detector = cv2.QRCodeDetector()

    data, bbox, straight_qrcode = detector.detectAndDecode(np_image)
    if bbox is not None:
        total = data[(data.find('s=') + 2):(data.find('&fn'))]
        dateD = data[(data.find('t=') + 8):(data.find('T'))]
        dateM = data[(data.find('t=') + 6):(data.find(dateD))]
        dateY = data[(data.find('t=') + 2):(data.find(dateM))]
        date = dateD + '.' + dateM + '.' + dateY
    else:
        total, date = None, None

    return total, date


def convert_image_to_numpy(filename):

    if filename.split('.')[-1] == 'pdf':
        pil_image = convert_from_path(Path(f'docs/{filename}'))[0]
        np_image = np.asarray(pil_image)
    else:
        pil_image = Image.open(f'docs/{filename}')
        np_image = np.asarray(pil_image)

    return np_image
