# AI Class Maker — PowerPoint API

FastAPI backend that turns a lesson topic, teacher instructions, and optional inspiration images into a downloadable editable PowerPoint (.pptx).

## Features
- AI-generated lesson slides for Grades 1–12
- Optional PNG, JPG, or WEBP inspiration images (up to 6)
- Teacher instructions, subject, language, and slide-count inputs
- Downloadable editable `.pptx` output
- Health endpoint for deployment checks

## Run locally
1. Install Python 3.10+.
2. Create a virtual environment and install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Set your API key:
   ```bash
   export OPENAI_API_KEY="your-api-key"
   ```
4. Start the API:
   ```bash
   uvicorn main:app --reload
   ```
5. Open http://localhost:8000/docs to test it.

## Deploy to Render
1. Create a new Web Service connected to this repository.
2. Render can use the included `render.yaml`, or set:
   - Build command: `pip install -r requirements.txt`
   - Start command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
3. Add `OPENAI_API_KEY` in Render → Environment. Never commit the real key.
4. Optional: set `OPENAI_MODEL` and `CORS_ORIGINS`.

The health endpoint is `GET /health`. It returns `ai_configured: true` when the API key is present.

## Generate a PowerPoint
Send a multipart/form-data POST to `/api/class-maker/generate` with:
- `topic` (required)
- `instructions`
- `grade` (1–12)
- `subject`
- `language`
- `slide_count` (3–20)
- `images` (optional; repeat the field for each image)

The response is a PowerPoint file with the MIME type
`application/vnd.openxmlformats-officedocument.presentationml.presentation`.

## Frontend integration
Use the endpoint:
`https://r-ai-pdf.onrender.com/api/class-maker/generate`

Build a `FormData` object and append the form fields and any selected images. Do not set the Content-Type header manually; the browser supplies the multipart boundary. The response is a binary PPTX, so read it as a Blob and download it with a temporary object URL.

Example:
```js
const form = new FormData();
form.append("topic", topic);
form.append("instructions", instructions);
form.append("grade", grade);
form.append("subject", subject);
form.append("language", language);
form.append("slide_count", "8");
for (const file of imageFiles) form.append("images", file);

const response = await fetch(
  "https://r-ai-pdf.onrender.com/api/class-maker/generate",
  { method: "POST", body: form }
);
if (!response.ok) {
  const error = await response.json().catch(() => ({}));
  throw new Error(error.detail || "Could not generate the PowerPoint.");
}
const blob = await response.blob();
const url = URL.createObjectURL(blob);
const a = document.createElement("a");
a.href = url;
a.download = "lesson.pptx";
a.click();
URL.revokeObjectURL(url);
```
