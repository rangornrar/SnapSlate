const SUPPORTED_LANGUAGES = new Set(["en", "fr"]);

const translations = {
  en: {
    meta: {
      title: "SnapSlate - Turn Windows captures into polished procedures | Archi-IT Labs",
      description:
        "SnapSlate turns Windows screenshots into polished procedures with OCR, annotations, project saving, export and publication-ready workflows.",
      keywords:
        "Windows screenshot annotation, procedure builder, OCR for screenshots, screenshot to document, technical documentation tool, support guide, QA documentation, Win Shift S",
      ogTitle: "SnapSlate - Turn Windows captures into polished procedures | Archi-IT Labs",
      ogDescription:
        "Capture screens, annotate them, keep steps organized and export polished procedures ready for documentation workflows.",
      ogImageAlt:
        "SnapSlate showing a polished procedure workspace with a canvas-first layout, floating tools and step panels.",
      twitterTitle: "SnapSlate - Turn Windows captures into polished procedures | Archi-IT Labs",
      twitterDescription:
        "Turn annotated Windows screenshots into exportable procedures, support guides and manuals.",
      locale: "en_US",
    },
    ui: {
      navLabel: "Primary navigation",
      languageSelector: "Language selector",
      languageLabel: "Language",
      previewHint: "Click to enlarge",
      lightboxTitle: "Enlarged preview",
      lightboxClose: "Close",
    },
    nav: {
      document: "Document",
      captures: "Screenshots",
      workflow: "Workflow",
      download: "Download",
      faq: "FAQ",
    },
    header: {
      kofi: "Ko-fi",
      repo: "GitHub",
    },
    hero: {
      badge: "Archi-IT Labs · V2026.04.29.001",
      title: "Turn Windows captures into polished procedures.",
      lead:
        "SnapSlate keeps the canvas central, the tools compact and the document ready to export.",
      primary: "Download installer",
      secondary: "Open GitHub",
      pillAria: "Highlights",
      pill1: "Canvas-first workspace",
      pill2: "OCR and annotations",
      pill3: "Projects and demos",
      pill4: "PNG · PDF · DOCX · Markdown · HTML",
      limit:
        "Current behavior: SnapSlate imports Win + Shift + S captures while the application is open.",
      panelAria: "SnapSlate procedure preview",
      panelAlt: "SnapSlate procedure workspace with canvas, floating tools and contextual panels",
      statsAria: "Quick summary",
      stat1: "canvas-first workspace",
      stat2: "featured screens",
      stat3: "annotation tools",
    },
    document: {
      eyebrow: "From capture to document",
      title: "Built for production, not for decoration.",
      lead:
        "SnapSlate keeps the canvas, steps and exports visible without turning the interface into a dashboard.",
      card1Tag: "01 · Capture",
      card1Title: "Start from the screen",
      card1Text:
        "Each screenshot becomes a step in the procedure. The real screen stays at the center of the workflow.",
      card2Tag: "02 · Explain",
      card2Title: "Add meaning",
      card2Text:
        "Use arrows, OCR, frames, stickers, masks and notes to make each step instantly readable.",
      card3Tag: "03 · Structure",
      card3Title: "Keep it tidy",
      card3Text:
        "Reorder steps, rename them and keep the document clear from first capture to final export.",
    },
    captures: {
      eyebrow: "Product shots",
      title: "The screenshots keep their real proportions.",
      lead:
        "Every capture opens larger with one click, and the site preserves the original aspect ratio instead of cropping the frame.",
      procedureAlt: "SnapSlate procedure screen",
      procedureTitle: "Procedure",
      procedureText:
        "Canvas-first, floating tools and contextual panels keep the workspace readable.",
      helpAlt: "SnapSlate help screen",
      helpTitle: "Capture / Help",
      helpText:
        "A compact help view for capture flows, OCR and keyboard shortcuts.",
      exportAlt: "SnapSlate export screen",
      exportTitle: "Export",
      exportText:
        "Outputs, destinations and document settings stay easy to find.",
      settingsAlt: "SnapSlate settings screen",
      settingsTitle: "Settings",
      settingsText:
        "Theme, language, capture behavior and defaults stay under control.",
    },
    workflow: {
      eyebrow: "Complete workflow",
      title: "Capture, annotate, add legends and publish HTML.",
      lead:
        "A four-step editorial flow, shown through full Lorem Ipsum screens.",
      storyTag: "Editorial sequence",
      storyTitle: "From raw screen to polished handoff",
      storyText:
        "The layout keeps one story per step, with screenshots shown at their natural size and a direct path to the final publication.",
      step1Tag: "01 · Capture",
      step1Title: "Capture",
      step1Text:
        "Start from a Lorem Ipsum screen and keep the original proportions intact.",
      step1Alt: "Lorem Ipsum capture screen",
      step2Tag: "02 · Annotate",
      step2Title: "Annotate",
      step2Text:
        "Add arrows, callouts, highlights, stickers and masks without losing the page context.",
      step2Alt: "Lorem Ipsum annotated screen",
      step3Tag: "03 · Legends",
      step3Title: "Legends",
      step3Text:
        "Write multi-line sticker legends in the properties panel and keep them readable.",
      step3Alt: "Lorem Ipsum legends screen",
      step4Tag: "04 · Publish HTML",
      step4Title: "HTML publication",
      step4Text:
        "Turn the procedure into a clean HTML page for sharing or reuse.",
      step4Alt: "Lorem Ipsum HTML publication screen",
    },
    faq: {
      eyebrow: "Before you install",
      title: "Short answers, no detours.",
      lead:
        "You can see how SnapSlate fits a production workflow.",
      q1: "Does SnapSlate support OCR?",
      a1:
        "Yes. OCR can add editable text to the current screen and to your notes.",
      q2: "Can I save and reopen a project?",
      a2:
        "Yes. Projects, demo content and document state can be saved and loaded.",
      q3: "Are publications included?",
      a3:
        "Yes. Publication flows cover Notion, Confluence and SharePoint targets.",
      q4: "Does the canvas stay dominant?",
      a4:
        "Yes. The layout was redesigned so the canvas stays at the center of the screen.",
    },
    cta: {
      eyebrow: "Ready to document properly?",
      title: "Download SnapSlate and turn your next captures into a usable manual.",
      lead:
        "Install the Windows app, capture a flow, add explanations, then export in the right format.",
      primary: "Download installer",
      secondary: "Open GitHub",
    },
    footer: {
      copy:
        "SnapSlate turns Windows screenshots into structured, exportable documentation. Part of Archi-IT Labs.",
      archi: "Archi-IT",
      wallcraft: "WallCraft",
      github: "GitHub",
      kofi: "Ko-fi",
    },
  },
  fr: {
    meta: {
      title: "SnapSlate - Transformer des captures Windows en procédures soignées | Archi-IT Labs",
      description:
        "SnapSlate transforme les captures Windows en procédures soignées avec OCR, annotations, sauvegarde du projet, export et publications prêtes à l'emploi.",
      keywords:
        "annotation capture Windows, générateur de procédures, OCR capture écran, capture vers document, outil documentation technique, guide support, documentation QA, Win Shift S",
      ogTitle: "SnapSlate - Transformer des captures Windows en procédures soignées | Archi-IT Labs",
      ogDescription:
        "Capturez des écrans, annotez-les, gardez vos étapes organisées et exportez des procédures prêtes pour vos workflows de documentation.",
      ogImageAlt:
        "SnapSlate montrant un espace de procédure soigné avec canvas central, outils flottants et panneaux d'étapes.",
      twitterTitle: "SnapSlate - Transformer des captures Windows en procédures soignées | Archi-IT Labs",
      twitterDescription:
        "Transformez des captures Windows annotées en procédures, guides support et manuels exportables.",
      locale: "fr_FR",
    },
    ui: {
      navLabel: "Navigation principale",
      languageSelector: "Sélecteur de langue",
      languageLabel: "Langue",
      previewHint: "Cliquer pour agrandir",
      lightboxTitle: "Aperçu agrandi",
      lightboxClose: "Fermer",
    },
    nav: {
      document: "Document",
      captures: "Captures",
      workflow: "Workflow",
      download: "Télécharger",
      faq: "FAQ",
    },
    header: {
      kofi: "Ko-fi",
      repo: "GitHub",
    },
    hero: {
      badge: "Archi-IT Labs · V2026.04.29.001",
      title: "Transformez des captures Windows en procédures soignées.",
      lead:
        "SnapSlate garde le canvas au centre, les outils compacts et le document prêt à exporter.",
      primary: "Télécharger l’installateur",
      secondary: "Ouvrir GitHub",
      pillAria: "Points forts",
      pill1: "Espace centré sur le canvas",
      pill2: "OCR et annotations",
      pill3: "Projets et démos",
      pill4: "PNG · PDF · DOCX · Markdown · HTML",
      limit:
        "Comportement actuel : SnapSlate importe les captures Win + Shift + S lorsque l’application est ouverte.",
      panelAria: "Aperçu de procédure SnapSlate",
      panelAlt: "Espace de procédure SnapSlate avec canvas, outils flottants et panneaux contextuels",
      statsAria: "Résumé rapide",
      stat1: "espace centré sur le canvas",
      stat2: "écrans clés",
      stat3: "outils d’annotation",
    },
    document: {
      eyebrow: "De la capture à la procédure",
      title: "Conçu pour produire, pas pour décorer.",
      lead:
        "SnapSlate garde visibles le canvas, les étapes et les exports sans transformer l’interface en tableau de bord.",
      card1Tag: "01 · Capturer",
      card1Title: "Partir de l’écran réel",
      card1Text:
        "Chaque capture devient une étape de la procédure. L’écran réel reste au centre du flux.",
      card2Tag: "02 · Annoter",
      card2Title: "Ajouter du sens",
      card2Text:
        "Utilisez flèches, OCR, cadres, gommettes, masques et notes pour rendre chaque étape immédiatement lisible.",
      card3Tag: "03 · Structurer",
      card3Title: "Rester lisible",
      card3Text:
        "Réordonnez les étapes, renommez-les et gardez un document clair de la première capture à l’export final.",
    },
    captures: {
      eyebrow: "Captures produit",
      title: "Les captures conservent leurs proportions réelles.",
      lead:
        "Chaque visuel s’ouvre en grand d’un clic, sans recadrage artificiel.",
      procedureAlt: "Écran Procédure de SnapSlate",
      procedureTitle: "Document de procédure",
      procedureText:
        "Canvas d’abord, outils flottants et panneaux contextuels gardent le workspace lisible.",
      helpAlt: "Écran Aide de SnapSlate",
      helpTitle: "Capturer / Aide",
      helpText:
        "Une vue compacte pour les flux de capture, l’OCR et les raccourcis clavier.",
      exportAlt: "Écran Export de SnapSlate",
      exportTitle: "Export",
      exportText:
        "Les sorties, destinations et réglages du document restent faciles à trouver.",
      settingsAlt: "Écran Réglages de SnapSlate",
      settingsTitle: "Réglages",
      settingsText:
        "Thème, langue, capture et paramètres par défaut restent sous contrôle.",
    },
    workflow: {
      eyebrow: "Workflow complet",
      title: "Capturer, annoter, ajouter les légendes et publier en HTML.",
      lead:
        "Un flux éditorial en quatre étapes, illustré par de vraies captures Lorem Ipsum.",
      storyTag: "Séquence éditoriale",
      storyTitle: "De l’écran brut à la livraison soignée",
      storyText:
        "La mise en page raconte une seule action par étape, avec des captures à leur taille naturelle et une sortie HTML finale claire.",
      step1Tag: "01 · Capturer",
      step1Title: "Capturer",
      step1Text:
        "Démarrez sur un écran Lorem Ipsum et gardez les proportions d’origine.",
      step1Alt: "Capture Lorem Ipsum",
      step2Tag: "02 · Annoter",
      step2Title: "Annoter",
      step2Text:
        "Ajoutez flèches, encadrés, surlignages, gommettes et masques sans perdre le contexte.",
      step2Alt: "Capture Lorem Ipsum annotée",
      step3Tag: "03 · Légendes",
      step3Title: "Légendes",
      step3Text:
        "Rédigez des légendes de gommettes multi-lignes dans le panneau de propriétés.",
      step3Alt: "Panneau des légendes Lorem Ipsum",
      step4Tag: "04 · Publication HTML",
      step4Title: "Publication HTML",
      step4Text:
        "Transformez la procédure en page HTML propre pour le partage ou la réutilisation.",
      step4Alt: "Publication HTML Lorem Ipsum",
    },
    faq: {
      eyebrow: "Avant d’installer",
      title: "Des réponses courtes, sans détour.",
      lead:
        "Vous voyez comment SnapSlate s’insère dans un vrai workflow de production.",
      q1: "SnapSlate gère-t-il l’OCR ?",
      a1:
        "Oui. L’OCR peut ajouter du texte modifiable sur l’écran courant et dans vos notes.",
      q2: "Peut-on sauvegarder et rouvrir un projet ?",
      a2:
        "Oui. Les projets, la démo et l’état du document peuvent être sauvegardés et rechargés.",
      q3: "Les publications sont-elles incluses ?",
      a3:
        "Oui. Les flux de publication couvrent Notion, Confluence et SharePoint.",
      q4: "Le canvas reste-t-il dominant ?",
      a4:
        "Oui. La mise en page a été repensée pour garder le canvas au centre de l’écran.",
    },
    cta: {
      eyebrow: "Prêt à documenter proprement ?",
      title: "Téléchargez SnapSlate et transformez vos prochaines captures en manuel utilisable.",
      lead:
        "Installez l’application Windows, capturez un parcours, ajoutez vos explications, puis exportez le document au bon format.",
      primary: "Télécharger l’installateur",
      secondary: "Ouvrir GitHub",
    },
    footer: {
      copy:
        "SnapSlate transforme des captures Windows en documentation structurée et exportable. Fait partie d’Archi-IT Labs.",
      archi: "Archi-IT",
      wallcraft: "WallCraft",
      github: "GitHub",
      kofi: "Ko-fi",
    },
  },
};

function normalizeLanguage(candidate) { if (!candidate) return null; const normalized = candidate.toLowerCase().slice(0, 2); return SUPPORTED_LANGUAGES.has(normalized) ? normalized : null; }
function detectBrowserLanguage() { const candidates = Array.isArray(navigator.languages) && navigator.languages.length > 0 ? navigator.languages : [navigator.language]; for (const candidate of candidates) { const normalized = normalizeLanguage(candidate); if (normalized) return normalized; } return "en"; }
function getInitialLanguage() { const url = new URL(window.location.href); const explicit = normalizeLanguage(url.searchParams.get("lang")); if (explicit) return explicit; return detectBrowserLanguage(); }
function setMeta(selector, value) { const element = document.querySelector(selector); if (element) element.setAttribute("content", value); }
function resolve(dictionary, key) { return key.split(".").reduce((value, part) => value?.[part], dictionary); }
function updateLanguageButtons(activeLanguage) { document.querySelectorAll("[data-lang-button]").forEach((button) => { const isActive = button.dataset.langButton === activeLanguage; button.classList.toggle("is-active", isActive); button.setAttribute("aria-pressed", String(isActive)); }); }
function applyLanguage(language, persist = true) { const dictionary = translations[language] ?? translations.en; document.documentElement.lang = language; document.title = dictionary.meta.title; setMeta('meta[name="description"]', dictionary.meta.description); setMeta('meta[name="keywords"]', dictionary.meta.keywords); setMeta('meta[property="og:locale"]', dictionary.meta.locale); setMeta('meta[property="og:title"]', dictionary.meta.ogTitle); setMeta('meta[property="og:description"]', dictionary.meta.ogDescription); setMeta('meta[property="og:image:alt"]', dictionary.meta.ogImageAlt); setMeta('meta[name="twitter:title"]', dictionary.meta.twitterTitle); setMeta('meta[name="twitter:description"]', dictionary.meta.twitterDescription); document.querySelectorAll("[data-i18n]").forEach((element) => { const translation = resolve(dictionary, element.dataset.i18n); if (element.hasAttribute("data-i18n-no-text")) { element.textContent = ""; return; } if (translation !== undefined && translation !== null && element.childElementCount === 0) element.textContent = translation; }); document.querySelectorAll("[data-i18n-attr]").forEach((element) => { const attribute = element.dataset.i18nAttr; const translation = resolve(dictionary, element.dataset.i18n); if (attribute && translation !== undefined && translation !== null) element.setAttribute(attribute, translation); }); updateLanguageButtons(language); if (persist) { const url = new URL(window.location.href); url.searchParams.set("lang", language); history.replaceState({}, "", url); } }
function wireLanguageButtons() { document.querySelectorAll("[data-lang-button]").forEach((button) => { button.addEventListener("click", () => { const language = normalizeLanguage(button.dataset.langButton) ?? "en"; applyLanguage(language, true); }); }); }
function wireLightbox() {
  const root = document.querySelector("[data-lightbox-root]");
  if (!root) return;
  const image = root.querySelector("[data-lightbox-image]");
  const caption = root.querySelector("[data-lightbox-caption]");
  const triggers = document.querySelectorAll("[data-lightbox-src]");
  let lastTrigger = null;

  const close = () => {
    root.hidden = true;
    document.body.classList.remove("lightbox-open");
    if (lastTrigger && typeof lastTrigger.focus === "function") lastTrigger.focus();
  };

  const open = (trigger) => {
    if (!(image instanceof HTMLImageElement)) return;
    const img = trigger.querySelector("img");
    const source = trigger.dataset.lightboxSrc || img?.getAttribute("src") || "";
    const label = img?.getAttribute("alt") || "";
    image.src = source;
    image.alt = label;
    if (caption) caption.textContent = label;
    root.hidden = false;
    document.body.classList.add("lightbox-open");
    const closeButton = root.querySelector(".lightbox__close");
    if (closeButton instanceof HTMLElement) closeButton.focus();
  };

  triggers.forEach((trigger) => {
    trigger.addEventListener("click", () => {
      lastTrigger = trigger;
      open(trigger);
    });
  });

  root.querySelectorAll("[data-lightbox-close]").forEach((button) => {
    button.addEventListener("click", close);
  });

  root.addEventListener("click", (event) => {
    if (event.target === root) close();
  });

  document.addEventListener("keydown", (event) => {
    if (!root.hidden && event.key === "Escape") close();
  });
}
document.addEventListener("DOMContentLoaded", () => { wireLanguageButtons(); wireLightbox(); applyLanguage(getInitialLanguage(), false); });
