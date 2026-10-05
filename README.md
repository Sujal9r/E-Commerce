# Zenvora

A responsive single-page shop prototype with a browser-based inventory panel.

## Run locally

```sh
npm install
npm run dev
```

## Production build

```sh
npm run build
```

Vite writes the deployable static site to `dist/`.

## Deploy to Vercel

The repository includes `vercel.json` with the Vite build and output settings. Push the project to the configured GitHub repository, import that repository in Vercel, and deploy. Vercel should use `npm install`, `npm run build`, and `dist` as the output directory.

## Demo limitations

Product, bag, and stock data are saved in the current browser with local storage. The project does not include a shared database, authenticated admin account, or payment processor. Product photography, fonts, and icons load from external services.

## Assignment documents

Run `npm run docs` to regenerate `Zenvora_Project_Report.pptx` and `Zenvora_Project_Report.pdf` from the screenshots in `assignment-assets/`.
