// Fonts bundled with the app, so nothing is requested from Google Fonts: metric-compatible
// substitutes for Office fonts used to lay out documents (Carlito ≈ Calibri, Caladea ≈
// Cambria, Arimo ≈ Arial, Tinos ≈ Times New Roman, Cousine ≈ Courier New) plus the UI fonts.
// Only the Latin subsets (Spanish and other Western languages) are included.
import.meta.glob(
  [
    '/node_modules/@fontsource/carlito/{latin,latin-ext}-{400,400-italic,700,700-italic}.css',
    '/node_modules/@fontsource/caladea/{latin,latin-ext}-{400,400-italic,700,700-italic}.css',
    '/node_modules/@fontsource/arimo/{latin,latin-ext}-{400,400-italic,600,600-italic,700,700-italic}.css',
    '/node_modules/@fontsource/tinos/{latin,latin-ext}-{400,400-italic,700,700-italic}.css',
    '/node_modules/@fontsource/cousine/{latin,latin-ext}-{400,400-italic,700,700-italic}.css',
    '/node_modules/@fontsource/inter/{latin,latin-ext}-{300,400,400-italic,500,600,600-italic,700}.css',
    '/node_modules/@fontsource/outfit/{latin,latin-ext}-{400,500,600,700,800}.css',
    '/node_modules/@fontsource/jetbrains-mono/{latin,latin-ext}-{400,500,600}.css',
  ],
  { eager: true },
);
