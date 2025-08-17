import json
import requests
from io import BytesIO
from PIL import Image
from datasets import load_dataset
from transformers import BlipProcessor, BlipForConditionalGeneration, Trainer, TrainingArguments
import torch

# -----------------------------
# Utilities
# -----------------------------
def has_valid_text_and_image(example):
    # Check if text exists and image URL exists
    return bool(example.get("text")) and bool(example.get("image"))

def preprocess_function(example, processor):
    input_text = example.get("text", "")
    image_url = example.get("image", "")

    if not input_text or not image_url:
        print(f"[SKIP] Missing text or image")
        return None

    try:
        resp = requests.get(image_url, timeout=10)
        resp.raise_for_status()
        image = Image.open(BytesIO(resp.content)).convert("RGB")
    except Exception as e:
        print(f"[SKIP] Failed to load image from {image_url}: {e}")
        return None

    # Format prompt
    prompt = f"Question: {input_text}\nAnswer:"

    try:
        inputs = processor(
            text=prompt,
            images=image,
            padding="max_length",
            truncation=True,
            return_tensors="pt",
        )
    except Exception as e:
        print(f"[ERROR] Processor failed - {e}")
        return None

    # If you don’t have assistant responses, use empty labels
    labels_text = ""
    labels = processor.tokenizer(
        labels_text,
        padding="max_length",
        truncation=True,
        return_tensors="pt",
    )

    labels_ids = labels.input_ids.squeeze(0)
    if processor.tokenizer.pad_token_id is None:
        processor.tokenizer.pad_token = processor.tokenizer.eos_token
    labels_ids[labels_ids == processor.tokenizer.pad_token_id] = -100

    return {
        "input_ids": inputs.input_ids.squeeze(0),
        "attention_mask": inputs.attention_mask.squeeze(0),
        "pixel_values": inputs.pixel_values.squeeze(0),
        "labels": labels_ids,
    }

# -----------------------------
# Load model + dataset
# -----------------------------
model_name = "Salesforce/blip-image-captioning-base"
processor = BlipProcessor.from_pretrained(model_name)
model = BlipForConditionalGeneration.from_pretrained(model_name)

dataset = load_dataset("json", data_files="blip_ready_final.jsonl", split="train")

# Filter only valid examples
dataset = dataset.filter(has_valid_text_and_image)

if len(dataset) == 0:
    raise ValueError("❌ Dataset is empty after filtering. Check JSONL format and image URLs.")

# Preprocess
columns_to_remove = dataset.column_names
dataset = dataset.map(
    lambda x: preprocess_function(x, processor),
    batched=False,
    remove_columns=columns_to_remove,
)

print("✅ Dataset ready, size:", len(dataset))

# -----------------------------
# Training setup
# -----------------------------
device = "cuda" if torch.cuda.is_available() else "cpu"
model.to(device)

training_args = TrainingArguments(
    output_dir="./blip-finetuned-biology",
    per_device_train_batch_size=2,
    gradient_accumulation_steps=4,
    num_train_epochs=3,
    logging_dir="./logs",
    logging_steps=10,
    save_steps=500,
    save_total_limit=2,
    remove_unused_columns=False,
    push_to_hub=False,
    report_to="none",
)

trainer = Trainer(
    model=model,
    args=training_args,
    train_dataset=dataset,
    tokenizer=processor.tokenizer,
)

trainer.train()

model.save_pretrained("./blip-finetuned-biology")
processor.save_pretrained("./blip-finetuned-biology")

print("🎉 Fine-tuning completed and model saved!")

# -----------------------------
# Inference test
# -----------------------------
def generate_answer(image_path, question, model, processor):
    image = Image.open(image_path).convert("RGB")
    inputs = processor(
        images=image,
        text=f"Question: {question}\nAnswer:",
        return_tensors="pt"
    ).to(model.device)

    out = model.generate(**inputs, max_new_tokens=50)
    return processor.decode(out[0], skip_special_tokens=True)

# Example:
# print(generate_answer("test.jpg", "What is in this image?", model, processor))
