/**
 * Generador de documents de Resolució de Convalidació de Crèdits
 * Versió per navegador (sense fs ni path — usa fetch + Blob + descàrrega automàtica)
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  ShadingType,
  ImageRun,
  Header,
  TextWrappingType,
  TextWrappingSide,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
} from "docx";

/**
 * Genera un document .docx de resolució de convalidació i el descarrega al navegador.
 *
 * @param {Object} dades - Dades variables del document
 * @param {string} dades.nomAlumne        - Nom complet de l'alumne/a
 * @param {string} dades.dni              - DNI de l'alumne/a
 * @param {string} dades.grau             - "superior" | "mitjà"
 * @param {string} dades.cicleCodi        - Codi del cicle formatiu (ex: "SAI0")
 * @param {string} dades.cicleNom         - Nom del cicle formatiu
 * @param {Array}  dades.moduls           - Mòduls convalidats [{codi, nom, nota}]
 * @param {string} dades.directora        - Nom de la directora
 * @param {string} dades.data             - Data de la resolució
 * @param {string} dades.ciutat           - Ciutat (per defecte "Barcelona")
 * @param {string} dades.nomCentre        - Nom del centre (línia gran)
 * @param {string} dades.subtitolCentre   - Subtítol del centre (línia petita)
 * @param {string} [outputFileName]       - Nom del fitxer de sortida (opcional)
 */
async function generarDocument(dades: any, outputFileName?: string) {
  const {
    nomAlumne,
    dni,
    grau = "superior",
    cicleNom,
    cicleCodi,
    moduls,
    directora,
    data,
    ciutat = "Barcelona",
    nomCentre = "Centre d\u2019Estudis Roca",
    subtitolCentre = "ESO \u2013 Batxillerat \u2013 Cicles Formatius",
  } = dades;

  // ---- Carreguem el logo via fetch (funciona al navegador) ----
  const logoResponse = await fetch("/logo.jpg");
  if (!logoResponse.ok) {
    throw new Error(`No s'ha pogut carregar el logo: ${logoResponse.statusText}`);
  }
  const logoData = await logoResponse.arrayBuffer();

  const cicleComplet = `${cicleNom} (${cicleCodi})`;
  const article = "l\u2019alumna"; // canvia a "l'alumne" si és masculí

  // Estils base
  const fontBase = { font: "Arial", size: 24 }; // 12pt
  const fontBold = { ...fontBase, bold: true };

  // Helpers del cos del document
  const buit = () =>
    new Paragraph({ children: [new TextRun({ ...fontBase, text: "" })] });

  const seccio = (text: string) =>
    new Paragraph({
      children: [new TextRun({ ...fontBold, text })],
    });

  const textJustificat = (parts: { text: string; bold?: boolean }[]) =>
    new Paragraph({
      alignment: AlignmentType.BOTH,
      children: parts.map(
        ({ text, bold }) => new TextRun({ ...fontBase, bold: !!bold, text })
      ),
    });

  const liniaMoldul = ({ codi, nom, nota }: { codi: string; nom: string; nota: number }) =>
    new Paragraph({
      children: [
        new TextRun({ ...fontBold, text: `${codi} \u2013 ${nom} \u2013 ` }),
        new TextRun({ ...fontBase, text: "Es trasllada la qualificaci\u00f3 obtinguda" }),
        new TextRun({ ...fontBold, text: ` ${nota}` }),
      ],
    });

  // ---- Capçalera amb logo i nom del centre ----
  const logoImg = new ImageRun({
    data: logoData,
    type: "jpg",
    transformation: {
      width: 83,
      height: 84,
    },
    floating: {
      horizontalPosition: {
        relative: HorizontalPositionRelativeFrom.COLUMN,
        offset: -234950,
      },
      verticalPosition: {
        relative: VerticalPositionRelativeFrom.PARAGRAPH,
        offset: -220980,
      },
      wrap: {
        type: TextWrappingType.SQUARE,
        side: TextWrappingSide.BOTH_SIDES,
      },
      allowOverlap: true,
      lockAnchor: false,
    },
  });

  const header = new Header({
    children: [
      new Paragraph({
        style: "Header",
        children: [
          logoImg,
          new TextRun({ text: " " }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: nomCentre,
            font: "Arial",
            size: 44,
            bold: true,
            color: "808080",
          }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: subtitolCentre,
            font: "Arial",
            size: 32,
            bold: true,
            color: "808080",
          }),
        ],
      }),
      new Paragraph({ children: [new TextRun({ text: "" })] }),
    ],
  });

  // ---- Cos del document ----
  const children = [
    buit(),
    buit(),

    new Paragraph({
      shading: { fill: "2F5496", type: ShadingType.CLEAR },
      children: [
        new TextRun({
          text: "Resoluci\u00f3 de convalidaci\u00f3 de cr\u00e8dits / m\u00f2duls / U.F. a Cicles Formatius ",
          font: "Arial",
          size: 26,
          bold: true,
          color: "FFFFFF",
        }),
      ],
    }),
    buit(),
    buit(),

    seccio("FETS"),
    buit(),
    textJustificat([
      { text: `Atesa la sol\u00b7licitud de convalidaci\u00f3 i la corresponent documentaci\u00f3 acreditativa presentada per ${article} ` },
      { text: `${nomAlumne} `, bold: true },
      { text: "amb DNI " },
      { text: `${dni} `, bold: true },
      { text: `matriculada en el cicle formatiu de grau ${grau} ` },
      { text: cicleComplet, bold: true },
      { text: "." },
    ]),
    buit(),
    buit(),

    seccio("FONAMENTS DE DRET"),
    buit(),
    textJustificat([
      { text: "At\u00e8s que la petici\u00f3 correspon als sup\u00f2sits previstos pel Servei d\u2019Organitzaci\u00f3 del Curr\u00edculum de la Formaci\u00f3 Professional Inicial, per la qual es determinen les convalidacions entre els m\u00f2duls establerts entre els diferents t\u00edtols de formaci\u00f3 professional." },
    ]),
    buit(),
    buit(),

    seccio("RESOLC"),
    buit(),
    textJustificat([
      { text: `Atorgar a ${article} ` },
      { text: `${nomAlumne} `, bold: true },
      { text: "la convalidaci\u00f3 dels seg\u00fcents cr\u00e8dits del cicle formatiu " },
      { text: cicleComplet, bold: true },
      { text: "." },
    ]),
    buit(),
    buit(),

    ...moduls.flatMap((m: { codi: string; nom: string; nota: number }) => [liniaMoldul(m), buit()]),

    buit(),
    buit(),

    new Paragraph({
      children: [
        new TextRun({ ...fontBase, text: "La Directora" }),
        new TextRun({ ...fontBase, text: "\t\t\t\t\t\t" }),
        new TextRun({ ...fontBase, text: `${ciutat}, ` }),
        new TextRun({ ...fontBold, text: data }),
      ],
    }),

    buit(),
    buit(),
    buit(),
    buit(),

    new Paragraph({
      children: [
        new TextRun({ ...fontBase, text: directora }),
        new TextRun({ ...fontBase, text: "\t\t\t\t\t\t" }),
        new TextRun({ ...fontBase, text: "Segell del centre" }),
      ],
    }),

    buit(),
  ];

  const doc = new Document({
    sections: [
      {
        headers: { default: header },
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1646, right: 1134, bottom: 1134, left: 1134 },
          },
        },
        children,
      },
    ],
  });

  // ---- Descàrrega al navegador (sense fs) ----
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download =
    outputFileName ||
    `RESOLUCIO_${nomAlumne.replace(/\s+/g, "_").toUpperCase()}.docx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  console.log(`✅ Document descarregat: ${a.download}`);
}

export { generarDocument };