## PPT Reconstruction & Auto-Styling Pipeline
### Overall Approach
The project converts an existing PowerPoint file into a clean, redesigned presentation automatically. The pipeline extracts slide text and positions, reconstructs logical slide structure (tables/bullets) using python, then refines the content using Gemini, and generates a new themed PowerPoint using python-pptx. 
Compatible for any number of slides and content type (tables, bullet points, paragraphs etc)
### How Scoring / Selection Works
Instead of a fixed scoring system, layout selection is dynamic and content-driven. The reconstructed slide grid is analyzed and Gemini determines the most suitable layout (table, bullet, or paragraph) based on the content structure. This adaptive approach works better for variable slide formats.
### Why the Template / Theme Was Selected
Gemini analyzes the presentation content and selects a consistent visual theme for all slides. The model generates a color palette including primary color, accent color, slide background, and text colors. This ensures visual consistency across the entire presentation.
### How to Run the Code
1. Install dependencies: pip install python-pptx google-generativeai
2. Add your Gemini API key, input ppt path and output file name in the script.
3. Run the Python script.
Example: final_function("input.pptx", "output.pptx", api_key)
### Libraries Used
- python-pptx – Reading slides, extracting shapes, and generating the final presentation.
- google-generativeai / google-genai – Refining text, choosing layout, and generating presentation theme.
- json – Used for structured data exchange between extraction, AI refinement, and slide generation.
