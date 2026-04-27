const SUPPORTED_LANGUAGES = new Set(["en", "fr"]);

const translations = {
  en: {
    meta: {
      title: "SnapSlate - Fast visual procedures for Windows | Archi-IT Labs",
      description:
        "SnapSlate turns Windows screenshots into clean, exportable procedures with annotations, sticker legends, and PNG, PDF, DOCX, Markdown, and HTML export.",
      ogTitle: "SnapSlate - Fast visual procedures for Windows | Archi-IT Labs",
      ogDescription:
        "Capture, annotate, and export Windows procedures with SnapSlate, an Archi-IT Labs tool.",
      ogImageAlt:
        "SnapSlate, a modern interface for turning a Windows capture into an exportable procedure.",
      twitterTitle: "SnapSlate - Fast visual procedures for Windows | Archi-IT Labs",
      twitterDescription:
        "Turn a Windows capture into a clean, exportable visual procedure.",
      locale: "en_US",
    },
    ui: {
      navLabel: "Primary navigation",
      languageSelector: "Language selector",
      languageLabel: "Language",
    },
    nav: {
      why: "Why",
      captures: "Screenshots",
      workflow: "Workflow",
      suite: "Suite",
      faq: "FAQ",
    },
    header: {
      kofi: "Ko-fi",
      repo: "Repository",
    },
    hero: {
      badge: "Archi-IT Labs · V2026.04.26.001",
      title: "Turn a Windows capture into a clean procedure.",
      lead:
        "SnapSlate follows Win + Shift + S, creates one step per capture, and keeps the canvas centered so you can annotate, label, and export without distraction.",
      primary: "Browse repository",
      secondary: "See how it works",
      kofi: "Support on Ko-fi",
      pillAria: "Highlights",
      pill1: "Win + Shift + S",
      pill2: "One step per capture",
      pill3: "Canvas first",
      pill4: "PNG, PDF, DOCX, Markdown, HTML",
      panelAria: "SnapSlate preview",
      panelAlt: "SnapSlate procedure view",
      statsAria: "Quick summary",
      stat1: "capture = 1 step",
      stat2: "visual palettes",
      stat3: "export formats",
    },
    why: {
      eyebrow: "Why SnapSlate",
      title: "A focused tool that stays fast without becoming heavy.",
      lead:
        "The goal is not to add an endless tool suite, but to save time on the real path: capture, explain, export.",
      card1Title: "Fast",
      card1Text:
        "A Windows capture becomes a new step almost immediately, without jumping to another tool.",
      card2Title: "Readable",
      card2Text:
        "The canvas stays central, the toolbar stays compact, and the property panel stays contextual.",
      card3Title: "Exportable",
      card3Text:
        "Captures become ready-made material for a manual, guide, ticket, or report.",
    },
    captures: {
      eyebrow: "Screenshots",
      title: "Real SnapSlate screenshots, captured from the app itself.",
      lead:
        "The page shows the software as it is actually used: procedure, help, export, and settings.",
      procedureAlt: "SnapSlate procedure view",
      procedureTitle: "Procedure",
      procedureText:
        "The main view keeps steps on the left, the canvas in the center, and the contextual panel on the right.",
      helpAlt: "SnapSlate capture help view",
      helpTitle: "Capture / Help",
      helpText:
        "A short, practical page that explains automatic import and the main shortcuts.",
      exportAlt: "SnapSlate export view",
      exportTitle: "Export",
      exportText:
        "Document settings, export folder, and formats stay in a clear, compact screen.",
      settingsAlt: "SnapSlate settings view",
      settingsTitle: "Settings",
      settingsText:
        "Theme, language, capture behavior, and default export stay visible without clutter.",
    },
    workflow: {
      eyebrow: "Workflow",
      title: "A short path: capture, annotate, export.",
      lead:
        "SnapSlate keeps the flow simple: the capture arrives, the procedure is built, and the guide is exported in the right format.",
      step1Title: "Capture",
      step1Text:
        "The user launches Win + Shift + S and SnapSlate watches the clipboard while the app is open.",
      step2Title: "Create a step",
      step2Text:
        "Each new capture becomes a step with its own title, note, and annotations.",
      step3Title: "Annotate",
      step3Text:
        "Tools stay compact, colors are quick to pick, and items remain editable after placement.",
      step4Title: "Organize",
      step4Text:
        "Steps can move, renumber, and keep a readable legend without taking over the layout.",
      step5Title: "Export",
      step5Text:
        "The document can go to PNG, PDF, DOCX, Markdown, or HTML depending on the audience.",
    },
    suite: {
      eyebrow: "Archi-IT Labs suite",
      title: "A coherent suite built around useful, sober tools.",
      lead:
        "The visual style, CTAs, and presentation logic follow the same family as the other Archi-IT Labs projects.",
      card1Title: "SnapSlate",
      card1Text: "Screenshots, procedures, annotations, and fast export for Windows.",
      card2Title: "WallCraft Pro",
      card2Text: "The visual companion in the suite for multi-monitor wallpapers.",
      card2Link: "See WallCraft",
      card3Title: "archi-it.fr",
      card3Text: "The main hub for projects, services, and references.",
      card3Link: "Open archi-it.fr",
      card4Title: "Ko-fi",
      card4Text: "A support link to keep the suite alive and keep shipping.",
      card4Link: "Support me",
    },
    faq: {
      eyebrow: "Before you download",
      title: "Short answers, no detours.",
      lead:
        "The goal is to understand at a glance how SnapSlate fits into a production workflow.",
      q1: "Does SnapSlate replace Win + Shift + S?",
      a1:
        "No. It relies on the Windows flow and watches the clipboard while it runs, so the capture stays familiar.",
      q2: "Which formats can I export?",
      a2:
        "The guide can go to PNG, PDF, DOCX, Markdown, or HTML depending on the documentation need.",
      q3: "Can steps be reordered?",
      a3:
        "Yes. Steps can be moved, reordered, and renumbered without losing the document logic.",
      q4: "Why Archi-IT Labs?",
      a4:
        "SnapSlate belongs to a suite of tools designed as one coherent set: same visual language, same product logic, same clarity.",
    },
    cta: {
      eyebrow: "Ready to try it?",
      title: "See the code, follow the updates, or support Archi-IT Labs.",
      lead:
        "SnapSlate is still a production tool. The site gives you the context, the screenshots, and the right links to integrate it into your suite.",
      primary: "GitHub",
      secondary: "Ko-fi",
    },
    footer: {
      copy:
        "SnapSlate · fast capture, annotation, and export. An Archi-IT Labs tool built to complement the suite.",
      archi: "Archi-IT",
      wallcraft: "WallCraft",
      github: "GitHub",
      kofi: "Ko-fi",
    },
  },
  fr: {
    meta: {
      title: "SnapSlate - Procédures visuelles rapides pour Windows | Archi-IT Labs",
      description:
        "SnapSlate transforme les captures Windows en procédures propres et exportables avec annotations, légendes de gommettes et export PNG, PDF, DOCX, Markdown et HTML.",
      ogTitle: "SnapSlate - Procédures visuelles rapides pour Windows | Archi-IT Labs",
      ogDescription:
        "Capture, annotation et export de procédures Windows avec SnapSlate, un outil Archi-IT Labs.",
      ogImageAlt:
        "SnapSlate, une interface moderne pour transformer une capture Windows en procédure exportable.",
      twitterTitle: "SnapSlate - Procédures visuelles rapides pour Windows | Archi-IT Labs",
      twitterDescription:
        "Transformez une capture Windows en procédure visuelle propre et exportable.",
      locale: "fr_FR",
    },
    ui: {
      navLabel: "Navigation principale",
      languageSelector: "Sélecteur de langue",
      languageLabel: "Langue",
    },
    nav: {
      why: "Pourquoi",
      captures: "Captures",
      workflow: "Workflow",
      suite: "Suite",
      faq: "FAQ",
    },
    header: {
      kofi: "Ko-fi",
      repo: "Dépôt",
    },
    hero: {
      badge: "Archi-IT Labs · V2026.04.26.001",
      title: "Transformez une capture Windows en procédure propre.",
      lead:
        "SnapSlate suit Win + Shift + S, crée une étape par capture et garde le canvas au centre pour annoter, légender et exporter sans détour.",
      primary: "Consulter le dépôt",
      secondary: "Voir le fonctionnement",
      kofi: "Me soutenir sur Ko-fi",
      pillAria: "Points forts",
      pill1: "Win + Shift + S",
      pill2: "Une étape par capture",
      pill3: "Canvas prioritaire",
      pill4: "PNG, PDF, DOCX, Markdown, HTML",
      panelAria: "Aperçu SnapSlate",
      panelAlt: "Vue Procédure de SnapSlate",
      statsAria: "Résumé rapide",
      stat1: "capture = 1 étape",
      stat2: "palettes visuelles",
      stat3: "formats d'export",
    },
    why: {
      eyebrow: "Pourquoi SnapSlate",
      title: "Un outil ciblé pour aller vite sans alourdir l’usage.",
      lead:
        "L'objectif n'est pas d'ajouter une suite sans fin de fonctions, mais de faire gagner du temps sur le vrai trajet: capturer, expliquer, exporter.",
      card1Title: "Rapide",
      card1Text:
        "Une capture Windows devient une nouvelle étape presque immédiatement, sans repasser par un logiciel intermédiaire.",
      card2Title: "Lisible",
      card2Text:
        "Le canvas reste au centre, la barre d'outils reste compacte et le panneau de propriétés reste contextuel.",
      card3Title: "Exportable",
      card3Text:
        "Les captures servent directement de matière à un manuel, un guide, un ticket ou un compte rendu.",
    },
    captures: {
      eyebrow: "Captures d'écran",
      title: "De vraies captures SnapSlate, prises dans l'application.",
      lead:
        "La page montre le logiciel tel qu’il est réellement utilisé: procédure, aide, export et réglages.",
      procedureAlt: "Vue Procédure de SnapSlate",
      procedureTitle: "Procédure",
      procedureText:
        "La vue principale concentre les étapes à gauche, le canvas au centre et le panneau contextuel à droite.",
      helpAlt: "Vue Aide de SnapSlate",
      helpTitle: "Capturer / Aide",
      helpText:
        "Une page courte et utile pour expliquer l'import automatique et les raccourcis principaux.",
      exportAlt: "Vue Export de SnapSlate",
      exportTitle: "Export",
      exportText:
        "Les paramètres du document, le dossier d’export et les formats restent dans un écran clair.",
      settingsAlt: "Vue Réglages de SnapSlate",
      settingsTitle: "Réglages",
      settingsText:
        "Thème, langue, capture et export par défaut restent visibles sans surcharge.",
    },
    workflow: {
      eyebrow: "Workflow",
      title: "Un trajet court: capturer, annoter, exporter.",
      lead:
        "SnapSlate suit une logique simple: la capture arrive, la procédure se construit, puis le guide part vers le bon format.",
      step1Title: "Capturer",
      step1Text:
        "L'utilisateur lance Win + Shift + S et SnapSlate surveille le presse-papiers quand l'application est ouverte.",
      step2Title: "Créer une étape",
      step2Text:
        "Chaque nouvelle capture devient une étape avec son titre, sa note et ses annotations propres.",
      step3Title: "Annoter",
      step3Text:
        "Les outils restent compacts, les couleurs sont rapides à choisir et les items restent modifiables après placement.",
      step4Title: "Organiser",
      step4Text:
        "Les étapes se déplacent, se renumérotent et gardent une légende lisible sans bloquer l'espace.",
      step5Title: "Exporter",
      step5Text:
        "Le document part en PNG, PDF, DOCX, Markdown ou HTML selon le besoin.",
    },
    suite: {
      eyebrow: "Suite Archi-IT Labs",
      title: "Une suite cohérente, pensée pour des outils utiles et sobres.",
      lead:
        "La charte, les CTA et la logique de présentation suivent la même famille visuelle que les autres projets Archi-IT Labs.",
      card1Title: "SnapSlate",
      card1Text: "Captures, procédures, annotations et export rapide pour Windows.",
      card2Title: "WallCraft Pro",
      card2Text: "Le compagnon visuel de la suite pour les fonds d'écran multi-écrans.",
      card2Link: "Voir WallCraft",
      card3Title: "archi-it.fr",
      card3Text: "Le hub principal pour retrouver les projets, les services et les références.",
      card3Link: "Ouvrir archi-it.fr",
      card4Title: "Ko-fi",
      card4Text: "Un lien de soutien pour garder la suite vivante et continuer les livraisons.",
      card4Link: "Me soutenir",
    },
    faq: {
      eyebrow: "Questions avant téléchargement",
      title: "Des réponses courtes, sans détour.",
      lead:
        "L'idée est de savoir en un coup d'œil comment SnapSlate s'insère dans un vrai usage de production.",
      q1: "SnapSlate remplace-t-il Win + Shift + S ?",
      a1:
        "Non. Il s'appuie sur le flux Windows et lit le presse-papiers pendant qu'il tourne, donc la capture reste familière.",
      q2: "Quels formats peut-on exporter ?",
      a2:
        "Le guide peut partir en PNG, PDF, DOCX, Markdown ou HTML selon le besoin documentaire.",
      q3: "Les étapes restent-elles réordonnables ?",
      a3:
        "Oui. Les étapes peuvent être déplacées, réordonnées et renumérotées sans perdre la logique du document.",
      q4: "Pourquoi Archi-IT Labs ?",
      a4:
        "SnapSlate s'inscrit dans une suite d'outils pensée comme un ensemble cohérent: mêmes codes visuels, même logique produit, même lisibilité.",
    },
    cta: {
      eyebrow: "Prêt à tester ?",
      title: "Voir le code, suivre les mises à jour, ou soutenir Archi-IT Labs.",
      lead:
        "SnapSlate reste un outil de production. Le site te donne le contexte, les captures et les bons liens pour l'intégrer dans ta suite.",
      primary: "GitHub",
      secondary: "Ko-fi",
    },
    footer: {
      copy:
        "SnapSlate · capture, annotation et export rapide. Un outil Archi-IT Labs pensé pour compléter la suite.",
      archi: "Archi-IT",
      wallcraft: "WallCraft",
      github: "GitHub",
      kofi: "Ko-fi",
    },
  },
};

function normalizeLanguage(candidate) {
  if (!candidate) {
    return null;
  }

  const normalized = candidate.toLowerCase().slice(0, 2);
  return SUPPORTED_LANGUAGES.has(normalized) ? normalized : null;
}

function detectBrowserLanguage() {
  const candidates = Array.isArray(navigator.languages) && navigator.languages.length > 0
    ? navigator.languages
    : [navigator.language];

  for (const candidate of candidates) {
    const normalized = normalizeLanguage(candidate);
    if (normalized) {
      return normalized;
    }
  }

  return "en";
}

function getInitialLanguage() {
  const url = new URL(window.location.href);
  const explicit = normalizeLanguage(url.searchParams.get("lang"));
  if (explicit) {
    return explicit;
  }

  return detectBrowserLanguage();
}

function setMeta(selector, value) {
  const element = document.querySelector(selector);
  if (element) {
    element.setAttribute("content", value);
  }
}

function resolve(dictionary, key) {
  return key.split(".").reduce((value, part) => value?.[part], dictionary);
}

function updateLanguageButtons(activeLanguage) {
  document.querySelectorAll("[data-lang-button]").forEach((button) => {
    const isActive = button.dataset.langButton === activeLanguage;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });
}

function applyLanguage(language, persist = true) {
  const dictionary = translations[language] ?? translations.en;

  document.documentElement.lang = language;
  document.title = dictionary.meta.title;
  setMeta('meta[name="description"]', dictionary.meta.description);
  setMeta('meta[property="og:locale"]', dictionary.meta.locale);
  setMeta('meta[property="og:title"]', dictionary.meta.ogTitle);
  setMeta('meta[property="og:description"]', dictionary.meta.ogDescription);
  setMeta('meta[property="og:image:alt"]', dictionary.meta.ogImageAlt);
  setMeta('meta[name="twitter:title"]', dictionary.meta.twitterTitle);
  setMeta('meta[name="twitter:description"]', dictionary.meta.twitterDescription);

  document.querySelectorAll("[data-i18n]").forEach((element) => {
    const translation = resolve(dictionary, element.dataset.i18n);
    if (translation !== undefined && translation !== null && element.childElementCount === 0) {
      element.textContent = translation;
    }
  });

  document.querySelectorAll("[data-i18n-attr]").forEach((element) => {
    const attribute = element.dataset.i18nAttr;
    const translation = resolve(dictionary, element.dataset.i18n);
    if (attribute && translation !== undefined && translation !== null) {
      element.setAttribute(attribute, translation);
    }
  });

  updateLanguageButtons(language);

  if (persist) {
    const url = new URL(window.location.href);
    url.searchParams.set("lang", language);
    history.replaceState({}, "", url);
  }
}

function wireLanguageButtons() {
  document.querySelectorAll("[data-lang-button]").forEach((button) => {
    button.addEventListener("click", () => {
      const language = normalizeLanguage(button.dataset.langButton) ?? "en";
      applyLanguage(language, true);
    });
  });
}

document.addEventListener("DOMContentLoaded", () => {
  wireLanguageButtons();
  applyLanguage(getInitialLanguage(), false);
});
