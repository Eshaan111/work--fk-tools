# Ecommerce Image Generation Workflow

When the user says, for example:

"10 White Baggy runs, use textured walls, don't use dark colors,
prefer studio lighting and professionalism"

follow this workflow automatically.

## Meaning of a run

- One run means one fresh background concept.
- Apply that same concept separately to every PNG in the relevant source folder.
- White Baggy currently contains five PNGs.
- Therefore, 10 runs means 10 concepts and 50 generated images.
- Use every available source PNG unless the user specifies otherwise.

## White Baggy mapping

- Source: `NO-BG-IMAGES/White-Baggy`
- Output: `IMAGES-FINAL/WHITE-BAGGY-JEANS`
- Excel category: `WHITE-BAGGY-JEANS`
- Vertical conventions may also be checked in `LEGACY/image_prompter.py`.

## Idea selection

- Read the `WHITE-BAGGY-JEANS` row from `USED-IMAGE-DESIGNS.xlsx`.
- Generate fresh concepts that do not repeat any exhausted idea.
- For each run, create one concise concept name.
- Append successfully completed concept names to the same Excel row.

## Generation requirements

Use the built-in ImageGen capability. No local prompting or generation script is required.

Treat every source PNG as an image-editing reference:

- Replace only the transparent background.
- Preserve the exact person, trousers, clothing, tattoos, accessories, pose,
  crop, body proportions, scale and camera perspective.
- Preserve the trousers' exact white colour, fabric texture, folds, seams,
  pockets and silhouette.
- Do not redraw or redesign the product.
- Do not add text, logos, watermarks, additional people or distracting objects.

When light professional textured backgrounds are requested:

- Use light or medium-light professional studio backgrounds.
- Use pale warm neutrals, pastels, soft blue, sage, peach, lavender,
  cream or similar colours.
- Use refined textures such as limewash, plaster, stucco, ribbing,
  microcement, linen render, clay or fine stone.
- Use clean ecommerce catalogue lighting, softbox illumination,
  realistic contact shadows and controlled highlights.
- Keep enough background contrast that white trousers remain clearly visible.
- Avoid black, charcoal, dark navy, deep brown, gloomy lighting,
  crushed shadows, clutter and dramatic cinematic scenes.

## Output structure

- Find the highest existing numeric folder in
  `IMAGES-FINAL/WHITE-BAGGY-JEANS`.
- Start with the next number.
- Create one numbered folder per concept.
- Preserve source mapping:
  - source `1.png` → output `1.png`
  - source `2.png` → output `2.png`
  - continue for every source PNG.

Example for 10 runs starting after folder 29:

`30/1.png` through `30/5.png`
...
`39/1.png` through `39/5.png`

## Validation

Before declaring completion:

- Confirm every numbered run folder exists.
- Confirm every run contains one result for every source PNG.
- Confirm all generated PNG files decode correctly.
- Visually inspect at least one output from every run.
- Update the Excel workbook only for successfully completed runs.
- Report the output folder range and total image count.
