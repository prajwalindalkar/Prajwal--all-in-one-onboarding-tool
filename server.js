'use strict';
const express = require('express');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, UnderlineType
} = require('docx');

const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── helpers ────────────────────────────────────────────────────────────────

const t = (text, opts = {}) => new TextRun({ text: String(text ?? ''), font: 'Arial', size: 22, ...opts });
const u = (text, opts = {}) => new TextRun({ text: String(text ?? ''), bold: true, underline: { type: UnderlineType.SINGLE }, font: 'Arial', size: 22, ...opts });
const chk = (checked) => new TextRun({ text: checked ? '☑' : '☐', font: 'Arial', size: 22 });
const p = (children, opts = {}) => new Paragraph({ children, spacing: { after: 80 }, ...opts });
const pBlank = () => new Paragraph({ children: [], spacing: { after: 80 } });

// Format ISO date YYYY-MM-DD → DD/MM/YYYY
function fmtDate(d) {
  if (!d) return '';
  const parts = d.split('-');
  if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
  return d;
}

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const allBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

function cell(children, opts = {}) {
  return new TableCell({
    borders: allBorders,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: Array.isArray(children) ? children : [p(children)],
    ...opts
  });
}

function headerCell(text, color = '5B2D8E') {
  return new TableCell({
    borders: allBorders,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    shading: { fill: color, type: ShadingType.CLEAR },
    children: [p([t(text, { bold: true, color: 'FFFFFF' })])],
    verticalAlign: VerticalAlign.CENTER
  });
}

const purpleRule = () => new Paragraph({
  children: [],
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: '5B2D8E', space: 1 } },
  spacing: { after: 80 }
});

function makeHeader(lh) {
  if (!lh || !lh.name) return undefined;
  const rightLines = [lh.address, lh.phone ? `Phone: ${lh.phone}` : null, lh.email ? `Email: ${lh.email}` : null, lh.website ? `Website: ${lh.website}` : null].filter(Boolean);
  return new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: lh.name, bold: true, font: 'Arial', size: 28 }),
          new TextRun({ text: '\t' + rightLines.join('  |  '), font: 'Arial', size: 18 })
        ],
        tabStops: [{ type: 'right', position: 9360 }],
        spacing: { after: 40 }
      }),
      purpleRule()
    ]
  });
}

// Build a 4-column signature table from a list of partners
function buildSignatureTable(presentPartners) {
  const rows = [];
  const chunks = [];
  for (let i = 0; i < presentPartners.length; i += 4) chunks.push(presentPartners.slice(i, i + 4));
  chunks.forEach(chunk => {
    while (chunk.length < 4) chunk.push(null);
    rows.push(new TableRow({ children: chunk.map(() => cell([p([t('____________')])])) }));
    rows.push(new TableRow({ children: chunk.map(pp => cell([p([pp ? u(pp.name) : t('')])])) }));
  });
  if (!rows.length) {
    rows.push(new TableRow({ children: [cell([p([t('____________')])]), cell([p([t('____________')])]), cell([p([t('')])]), cell([p([t('')])])] }));
    rows.push(new TableRow({ children: [cell([p([t('')])]), cell([p([t('')])]), cell([p([t('')])]), cell([p([t('')])])] }));
  }
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [2257, 2257, 2256, 2256], rows });
}

// ─── MDF Generator ──────────────────────────────────────────────────────────
// NO letterhead for MDF. Date field added.
function generateMDF(d) {
  const { business, as: as_ } = d;
  const sameOffice = business.principalSame === true || business.principalSame === 'true';

  const children = [
    p([t('Merchant Declaration', { bold: true, size: 28 })], { alignment: AlignmentType.CENTER }),
    pBlank(),
    p([t('To,')]),
    p([t('PhonePe Limited (Formerly known as \'PhonePe Private Limited\') (hereinafter referred as "PhonePe")')]),
    p([t('Office-2, Floor 5, Wing A, Block A, Salarpuria Softzone,')]),
    p([t('Bellandur Village, Varthur Hobli, Outer Ring Road, Bangalore South,')]),
    p([t('Bangalore, Karnataka, India, 560103')]),
    pBlank(),
    p([t('Subject: PhonePe Merchant Declaration')]),
    pBlank(),
    p([
      t('I, '), u(as_.name), t(', hereinafter referred to as '),
      t('"Merchant"', { bold: true }), t(' being the, '), u(as_.designation),
      t(' <Owner / Karta / Partner/ Director / Managing Director / Authorised Signatory> of the '),
      u(business.name), t(' <Firm name / Company name> having its <registered office> address at '),
      u(business.registeredAddress), t(' ("Entity") and having its <principal place of operation/office> at '),
      sameOffice ? u('same ✓') : u(business.principalAddress),
      t(', do hereby declare that I have been authorised, to act as a designated authorised signatory for the Entity (including, but not limited to, registration/execution/renewal/amendment of the business related association(s)/partnership(s)/contract(s)/terms and conditions with PhonePe) and that the below mentioned details provided by me (including my specimen signature) are true, accurate, valid, legally binding and authenticated for the Entity, and can be used the purposes of obtaining payment facilitation services, business related associations / partnership(s) with PhonePe.'),
    ]),
    pBlank(),
    p([t('I hereby allow PhonePe, to collect, store and use my KYC and/or other details as required by PhonePe, for the purposes of verifying my identity as the authorised signatory of the entity thereby, enabling the entity, to be onboarded as Merchant with PhonePe for the purposes of availing PhonePe services, in accordance with PhonePe\'s Terms and Conditions and Privacy Policy.')]),
    pBlank(),
    p([t('Details provided under this declaration:', { bold: true })]),
    pBlank(),
    p([t('1.   Mobile No. (registered with PhonePe for Onboarding): '), u(as_.mobile)]),
    p([t('2.   Email ID (registered with PhonePe for Onboarding): '), u(as_.email)]),
    p([t('3.   Individual KYC Documents:')]),
    p([t('4.   In case of Person with Disability (PwD), please specify')]),
    p([t('     Type of Disability:          '), u('NA')]),
    p([t('     Percentage of Disability:    '), u('NA')]),
    p([t('5.   Father\'s Name (of the Authorized Signatory):  '), u(as_.fatherName)]),
    pBlank(),
    p([t('PAN CARD (Mandatory)   '), chk(true)]),
    pBlank(),
    p([t('Any one of the following is mandatory (Please tick whichever submitted):', { italics: true })]),
    p([t('  · Aadhaar (masked except the last 4 digits)   '), chk(as_.ovdType === 'Aadhaar')]),
    p([t('  · Driving License                              '), chk(as_.ovdType === 'Driving License')]),
    p([t('  · Voter ID                                     '), chk(as_.ovdType === 'Voter ID')]),
    pBlank(),
    p([t('I hereby declare that the above information/ details provided herein are true, valid and accurate as on date of submission and further that I would be liable for any incorrect/false information or for any untrue statement of details / information provided.')]),
    pBlank(),
    new Table({
      width: { size: 9026, type: WidthType.DXA }, columnWidths: [4513, 4513],
      rows: [
        new TableRow({ children: [cell([p([t('Signature with seal:', { bold: true })])]), cell([p([t('')])])] }),
        new TableRow({ children: [cell([p([t('Name: '), u(as_.name)])]), cell([p([t('')])])] }),
      ]
    }),
    pBlank(),
    p([t('I, on behalf of the Merchant, further declare that:')]),
    pBlank(),
    new Table({
      width: { size: 9026, type: WidthType.DXA }, columnWidths: [7526, 1500],
      rows: [
        new TableRow({
          children: [
            cell([
              p([t('The below in the entity duly registered under the applicable law of lands:', { bold: true })]),
              p([chk(business.entityType === 'Individual/Sole Proprietorship'), t('  Individual/Sole Proprietorship          '), chk(business.entityType === 'HUF'), t('  HUF')]),
              p([chk(business.entityType === 'Company'), t('  Company                                  '), chk(business.entityType === 'Trust'), t('  Trust')]),
              p([chk(business.entityType === 'Registered Partnership Firm'), t('  Registered Partnership Firm             '), chk(business.entityType === 'Society'), t('  Society')]),
              p([chk(business.entityType === 'Un-Registered Partnership Firm'), t('  Un-Registered Partnership Firm          '), chk(business.entityType === 'Others'), t('  Others (Club, University, Institution etc.)')]),
              p([chk(business.entityType === 'Body of individual'), t('  Body of individual                       '), chk(business.entityType === 'Association of Person'), t('  Association of Person')]),
              p([chk(false), t('  Government Departments / Public Sector Undertaking / Local Government Bodies')]),
            ]),
            cell([p([t('Tick (as applicable', { italics: true })])])
          ]
        }),
       new TableRow({
  children: [
    cell([p([
      t('1. The Merchant is registered under Income Tax Act, 1961 (as may be amended from time to time) and has obtained TAN Number '),
      u(as_.tan || '__________________________'), t(' against the registration. '), t('OR', { bold: true })
    ])]),
    cell([p([t(as_.tan ? 'Tick (if applicable' : '', { italics: true })])])
  ]
}),
new TableRow({
  children: [
    cell([p([
      t('The Merchant does not hold TAN as it is not liable to deduct tax at source or collect tax at source as per the provisions of Income Tax Act, 1961.')
    ])]),
    cell([p([chk(!as_.tan)])])
  ]
}),
new TableRow({
  children: [
    cell([p([
      t('2. The Merchant is registered and a GSTIN certificate/acknowledgement having provisional number '),
      u(as_.gstin || '__________________________'), t(' is issued by GST authorities. '), t('OR', { bold: true })
    ])]),
    cell([p([t(as_.gstin ? 'Tick (if applicable' : '', { italics: true })])])
  ]
}),
new TableRow({
  children: [
    cell([p([t('The Merchant does not have any registration with GST authorities.')])]),
    cell([p([chk(!as_.gstin)])])
  ]
}),
        new TableRow({
          children: [
            cell([p([t('The entity is working in the nature of:')]),
              p([chk(business.nature === 'Government'), t('  Government organization   '), chk(business.nature === 'NGO'), t('  NGO/Charitable institution   '), chk(business.nature === 'NA' || !business.nature), t('  NA')])]),
            cell([p([t('Tick (as applicable', { italics: true })])])
          ]
        }),
        new TableRow({
          children: [
            cell([p([t('No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a Politically Exposed Person (PEP) (as defined by RBI).')])]),
            cell([p([chk(true)])])
          ]
        }),
      ]
    }),
    pBlank(),
    p([t('I, having PAN number '), u(as_.pan), t(', hereby declare that the above facts and information are true, complete and correct to the best of my knowledge. I understand and agree that in case it is found that the above-mentioned facts and information are incorrect, I will be personally held liable for the same.')]),
    pBlank(),
    new Table({
      width: { size: 9026, type: WidthType.DXA }, columnWidths: [4513, 4513],
      rows: [
        new TableRow({
          children: [
            cell([
              p([t('Yours faithfully,')]),
              p([t('For and behalf of the Merchant')]),
              pBlank(),
              p([t('_______________________')]),
              p([t('(Signature with Seal)')]),
              p([t('Designation: '), u(as_.designation)]),
              p([t('Date: '), u(fmtDate(as_.date))]),
              p([t('Place: '), u(business.city || '')]),
            ]),
            cell([p([t('Picture of the Authorised Signatory (Countersign with face visible)', { italics: true })])])
          ]
        })
      ]
    }),
    pBlank(),
    p([t('Note:', { bold: true })]),
    p([t('"Government company" means any company in which not less than fifty-one per cent of the paid-up share capital is held by the Central Government, or by any State Government or Governments.', { italics: true })]),
    p([t('NGO (For Darpan applicability): For Company – If registered as section 8 company. For Societies and Trust – all non-government societies and trusts.', { italics: true })]),
  ];

  // No letterhead for MDF
  return new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      children
    }]
  });
}

// ─── Partner Resolution Generator ───────────────────────────────────────────
function generateResolution(d) {
  const { business, partners, resolution, lh } = d;
  const sameOffice = business.principalSame === true || business.principalSame === 'true';

  // presentIds may be strings or numbers — normalise to strings
  const presentIds = (resolution.presentIds || []).map(String);
  const presentPartners = partners.filter(pp => presentIds.includes(String(pp.id)));
  const authPartner = partners.find(pp => pp.isAuth);

  // FIX: registered/unregistered rendering — only show the relevant checkbox+question
  const isRegistered = business.firmType === 'Registered';
  const isUnregistered = business.firmType === 'Unregistered';

  const declarationItems = [];
  // Registered block
  declarationItems.push(p([chk(isRegistered), t('  Registered')]));
  if (isRegistered) {
    declarationItems.push(p([t('   Whether there is/are any natural person(s), acting alone or through other juridical persons, with more than 10% ownership/entitlement to capital/profits, or who exercise control through other means.')]));
    declarationItems.push(p([t('   '), chk(business.naturalPersons10 === 'Yes'), t('  Yes      '), chk(business.naturalPersons10 === 'No'), t('  No')]));
    declarationItems.push(pBlank());
  }
  // Unregistered block
  declarationItems.push(p([chk(isUnregistered), t('  Unregistered')]));
  if (isUnregistered) {
    declarationItems.push(p([t('   Whether there is/are any natural person(s), acting alone or through other juridical persons, with more than 15% ownership/entitlement to capital/profits, or who exercise control through other means.')]));
    declarationItems.push(p([t('   '), chk(business.naturalPersons15 === 'Yes'), t('  Yes      '), chk(business.naturalPersons15 === 'No'), t('  No')]));
    declarationItems.push(pBlank());
  }

  const sigTable = buildSignatureTable(presentPartners);

  const children = [
    p([t('(on partnership firm\'s letterhead)', { italics: true })], { alignment: AlignmentType.CENTER }),
    p([t('TO WHOMSOEVER IT MAY CONCERN', { bold: true })], { alignment: AlignmentType.CENTER }),
    pBlank(),
    p([
      t('RESOLUTION OF THE PARTNERS PASSED AT THE MEETING OF THE PARTNERS OF ', { bold: true }),
      u(business.name, { bold: true }),
      t(' ("FIRM") HELD ON ', { bold: true }),
      u(fmtDate(resolution.date), { bold: true }),
      t(' AT ', { bold: true }),
      u(resolution.time, { bold: true }),
      t(' having its <registered office address> at '),
      u(business.registeredAddress),
      t(' and having its <principal place of operation/office> at '),
      sameOffice ? u('same ✓') : u(business.principalAddress),
      t(','),
    ]),
    pBlank(),
    p([t('PRESENT:', { bold: true })]),
    // FIX: show ALL partners, present ones with their name filled in bold+underline
    ...partners.map((pp, i) => {
      const isPresent = presentIds.includes(String(pp.id));
      return p([t(`${i + 1}. `), isPresent ? u(pp.name) : t(`_______________________`)]);
    }),
    p([t('(List of partner present during the resolution)', { italics: true })]),
    pBlank(),
    p([
      t('RESOLVED THAT', { bold: true }),
      t(' Mr/Mrs '), u(authPartner ? authPartner.name : '__________'),
      t(', '), t('<Partner>', { italics: true }),
      t(', be and is hereby authorized, to act on behalf of the Firm and to execute/sign all necessary applications/ documents for the purpose of opening and operating a business account with PhonePe Limited.'),
    ]),
    pBlank(),
    p([
      t('RESOLVED FURTHER THAT', { bold: true }),
      t(' all acts, deeds, and things done by the said Partner(s) in this regard shall be binding upon the Firm and all its partners and shall remain in force. '),
      t('RESOLVED FURTHER THAT', { bold: true }),
      t(' this resolution shall remain in force until a written notice of its withdrawal or amendment is served upon and acknowledged by PhonePe Limited, and that a certified true copy of this resolution be furnished to PhonePe Limited for their records'),
    ]),
    pBlank(),
    sigTable,   // FIX: names appear under signatures
    pBlank(),
    p([t('Seal of the Firm', { italics: true })], { alignment: AlignmentType.RIGHT }),
    pBlank(),
    p([t('Declaration', { bold: true, underline: { type: UnderlineType.SINGLE } })], { alignment: AlignmentType.CENTER }),
    pBlank(),
    p([t('I/we, the undersigned individuals, hereby personally, jointly, and severally undertake and declare that:')]),
    pBlank(),
    p([t('1. Our firm '), u(business.name), t(' is constituted as a partnership firm and it is')]),
    ...declarationItems,  // FIX: only relevant firm type + answer shown
    p([t('2. If there is/are no natural person(s) as per the responses provided under declaration (1) hereinabove, then we declare that the authorised signatory identified in the Resolution is the senior managing official and be considered as the beneficial owner of our firm.')]),
    pBlank(),
    p([t('3. Our personnel(s), partner(s), director(s), officer(s), or our family member(s) or our close associate(s) and beneficial owners, is a Politically Exposed Person. ("Politically Exposed Persons" (PEPs) are individuals who are or have been entrusted with prominent public functions by a foreign country, including the Heads of States/Governments, senior politicians, senior government or judicial or military officers, senior executives of state-owned corporations and important political party officials.)')]),
    p([t('   '), chk(business.pep === 'YES'), t('  YES      '), chk(business.pep === 'NO' || !business.pep), t('  NO')]),
    pBlank(),
    p([t('4. The contents of the resolution of the partners provided to PhonePe are true and valid.')]),
    p([t('5. The list of partners constituting the partnership firm and their respective details, as provided in the partnership deed, provided to PhonePe, are true, complete, current, and valid.')]),
    p([t('6. The percentage of ownership and/or entitlement to capital or profits of the partners as specified in the Partnership Deed is true, complete, current, and valid.')]),
    pBlank(),
    p([t('In consideration of PhonePe agreeing to rely on the Resolution, the Partnership Deed, and this declaration, I/we hereby personally, jointly, and severally undertake to indemnify and hold PhonePe harmless against all damages, liabilities, claims, demands, actions, proceedings, losses, costs (including legal costs), expenses, and all other liabilities of whatsoever nature or description, arising out of or in connection with PhonePe\'s reliance on the said Resolution, Partnership Deed, and this declaration.')]),
    pBlank(),
    sigTable,   // FIX: second signature table at end of declaration also has names
    // FIX: whitener line removed
  ];

  const header = makeHeader(lh);
  return new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      ...(header ? { headers: { default: header } } : {}),
      children
    }]
  });
}

// ─── Partnership BO Declaration Generator ───────────────────────────────────
function generatePartnershipBO(d) {
  const { business, partners, boDeclaration, lh } = d;
  const sameOffice = business.principalSame === true || business.principalSame === 'true';
  const authPartner = partners.find(pp => pp.isAuth);
  const threshold = business.firmType === 'Registered' ? 10 : 15;
  const boPartners = partners.filter(pp => parseFloat(pp.share) > threshold);

  const boTableRows = [
    new TableRow({
      children: [
        headerCell('S.N.'), headerCell('Name'), headerCell('Residential Address & PIN code'),
        headerCell('Designation'), headerCell('DOB'), headerCell('Proof of Identity (PAN)'),
        headerCell('Proof of Address'), headerCell('Nationality'), headerCell('% of interest/ownership')
      ]
    }),
    ...boPartners.map((pp, i) => new TableRow({
      children: [
        cell([p([u(`${i + 1}`)])]),
        cell([p([u(pp.name)])]),
        cell([p([u(pp.address)])]),
        cell([p([u(pp.designation || 'Partner')])]),
        cell([p([u(fmtDate(pp.dob))])]),
        cell([p([u(pp.pan)])]),
        cell([p([u(pp.ovdType === 'Aadhaar' ? `XXXX XXXX ${pp.aadhaarLast4}` : pp.ovdNumber || '')])]),
        cell([p([u(pp.nationality || 'Indian')])]),
        cell([p([u(`${pp.share}%`)])]),
      ]
    })),
    ...Array(Math.max(0, 4 - boPartners.length)).fill(null).map(() =>
      new TableRow({ children: Array(9).fill(null).map(() => cell([p([t('')])])) })
    )
  ];

  const children = [
    p([t('DECLARATION OF BENEFICIAL OWNERSHIP (BO) and LIST OF PARTNERS', { bold: true, underline: { type: UnderlineType.SINGLE } })], { alignment: AlignmentType.CENTER }),
    p([t('(Not applicable for Individual/ Sole Proprietor/ HUF/ Government Departments / Public Sector Undertaking / Local Authority)', { italics: true, size: 18 })], { alignment: AlignmentType.CENTER }),
    pBlank(),
    new Table({
      width: { size: 9026, type: WidthType.DXA }, columnWidths: [500, 3000, 5526],
      rows: [
        new TableRow({ children: [cell([p([t('I')])]), cell([p([t('Name of the entity')])]), cell([p([u(business.name)])])] }),
        new TableRow({ children: [cell([p([t('II')])]), cell([p([t('Registered address')])]), cell([p([u(business.registeredAddress)])])] }),
        new TableRow({ children: [cell([p([t('III')])]), cell([p([t('Principal place of operation/office')])]), cell([p([sameOffice ? u('✓ Same') : u(business.principalAddress)])])] }),
        new TableRow({ children: [cell([p([t('IV')])]), cell([p([t('Type of entity')])]), cell([p([chk(business.firmType === 'Registered'), t('  Partnership Firm / LLP   '), chk(business.firmType === 'Unregistered'), t('  Unregistered Partnership Firm')])])] }),
      ]
    }),
    pBlank(),
    p([t('The Legal Entity as stated above hereby confirms and declares the following on the below date: '), u(fmtDate(boDeclaration.date))]),
    pBlank(),
    p([t('Tick the box as applicable (Refer note A "RBI guidelines for identification of Beneficial owners")', { italics: true })]),
    pBlank(),
    p([chk(boDeclaration.category === '1'), t('  ', { bold: true }), t('Category 1', { bold: true }), t(' - We hereby declare that following persons/ entity noted in the below table own 10%/15% or more interest or possess the right to control management/policy decisions (Refer Notes).')]),
    p([chk(boDeclaration.category === '2'), t('  ', { bold: true }), t('Category 2', { bold: true }), t(' - We hereby declare that no natural person is identified as per category 1 (above). (Mention the details of the natural person(s) holding the position of senior management official in the entity.)')]),
    pBlank(),
    p([chk(!business.pep || business.pep === 'NO'), t('  No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a '), t('Politically Exposed Person (PEP)', { bold: true }), t(' (as defined by RBI).')]),
    pBlank(),
    p([t('The details of beneficial owner(s) is/are as follows:')]),
    pBlank(),
    new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [300, 900, 1400, 700, 700, 900, 1300, 700, 1126], rows: boTableRows }),
    pBlank(),
    p([t('List of the Current Partners of the firm operating at the aforementioned address:')]),
    ...partners.map((pp, i) => p([t(`${i + 1}. `), u(pp.name)])),
    pBlank(),
    p([t('We acknowledge and confirm that the information provided above is true and correct to the best of our knowledge and belief.')]),
    pBlank(),
    p([t('Authorised Signatory/ies: '), t('(Refer note B for signature requirement)', { italics: true })]),
    pBlank(),
    p([u(authPartner ? authPartner.name : '__________'), t('  (Name, Signature with Stamp)')]),
    pBlank(),
    p([t('Firm Seal is mandatory', { bold: true })]),
    pBlank(),
    new Paragraph({ children: [t('─'.repeat(80))], spacing: { after: 80 } }),
    pBlank(),
    p([t('#Notes:', { bold: true })]),
    p([t('A. RBI guidelines for identification of Beneficial owners', { bold: true })]),
    p([t('   Category 1: Controlling ownership interest means:', { bold: true })]),
    new Table({
      width: { size: 5000, type: WidthType.DXA }, columnWidths: [2500, 2500],
      rows: [
        new TableRow({ children: [headerCell('Business entity'), headerCell('Shareholding %')] }),
        new TableRow({ children: [cell([p([t('Partnership Firm')])]), cell([p([t('>10%')])])] }),
        new TableRow({ children: [cell([p([t('Unregistered Partnership Firm')])]), cell([p([t('>15%')])])] }),
      ]
    }),
    pBlank(),
    p([t('B. Signature on the Declaration form: A person who is authorised to sign BO declaration: CS / Authorised signatory', { bold: true })]),
    pBlank(),
    p([t('C. Other Instructions', { bold: true })]),
    p([t('   1. Proof of Identity – Individual: PAN*')]),
    p([t('   2. Proof of Address – Individual: Voter ID / Driving License / Passport / Redacted Aadhaar')]),
    p([t('   3. PAN Number to be provided for Residents/Entities registered in India.')]),
  ];

  const header = makeHeader(lh);
  return new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      ...(header ? { headers: { default: header } } : {}),
      children
    }]
  });
}

// ─── Company BO Declaration Generator ───────────────────────────────────────
function generateCompanyBO(d) {
  const { business, directors, companyBO, lh } = d;
  const sameOffice = business.principalSame === true || business.principalSame === 'true';

  const boTableRows = [
    new TableRow({
      children: [
        headerCell('S.N.'), headerCell('Name'), headerCell('Residential Address & PIN code'),
        headerCell('Designation'), headerCell('DOB'), headerCell('Proof of Identity (PAN)'),
        headerCell('Proof of Address'), headerCell('Nationality'), headerCell('% of interest/ownership')
      ]
    }),
    ...directors.map((dir, i) => new TableRow({
      children: [
        cell([p([u(`${i + 1}`)])]),
        cell([p([u(dir.name)])]),
        cell([p([u(dir.address)])]),
        cell([p([u(dir.designation)])]),
        cell([p([u(fmtDate(dir.dob))])]),
        cell([p([u(dir.pan)])]),
        cell([p([u(dir.ovdType === 'Aadhaar' ? `XXXX XXXX ${dir.aadhaarLast4}` : dir.ovdNumber || '')])]),
        cell([p([u(dir.nationality || 'Indian')])]),
        cell([p([u(`${dir.share}%`)])]),
      ]
    })),
    ...Array(Math.max(0, 5 - directors.length)).fill(null).map(() =>
      new TableRow({ children: Array(9).fill(null).map(() => cell([p([t('')])])) })
    )
  ];

  const mgmtRows = [
    new TableRow({ children: [headerCell('Name'), headerCell('Designation')] }),
    ...(companyBO.seniorMgmt || []).map(m => new TableRow({
      children: [cell([p([u(m.name)])]), cell([p([u(m.designation)])])]
    })),
    ...Array(Math.max(0, 3 - (companyBO.seniorMgmt || []).length)).fill(null).map(() =>
      new TableRow({ children: [cell([p([t('')])]), cell([p([t('')])])] })
    )
  ];

  const children = [
    p([t('Declarations', { bold: true, underline: { type: UnderlineType.SINGLE } })], { alignment: AlignmentType.CENTER }),
    pBlank(),
    p([t('A. DECLARATION OF BENEFICIAL OWNERSHIP (BO)', { bold: true })]),
    pBlank(),
    new Table({
      width: { size: 9026, type: WidthType.DXA }, columnWidths: [500, 2500, 6026],
      rows: [
        new TableRow({ children: [cell([p([t('I')])]), cell([p([t('Name of the entity')])]), cell([p([u(business.name)])])] }),
        new TableRow({ children: [cell([p([t('II')])]), cell([p([t('Registered address')])]), cell([p([u(business.registeredAddress)])])] }),
        new TableRow({ children: [cell([p([t('III')])]), cell([p([t('Principal place of operation/office')])]), cell([p([sameOffice ? u('✓ Same as above') : u(business.principalAddress)])])] }),
        new TableRow({ children: [cell([p([t('IV')])]), cell([p([t('Type of entity')])]), cell([p([u('Company')])])] }),
        new TableRow({
          children: [
            cell([p([t('V')])]),
            cell([p([t('A. Please tick the box as applicable –')])]),
            cell([
              p([chk(companyBO.listingType === 'i'), t('  i. An entity listed on a stock exchange in India, If yes, mention the Stock Exchange Name: '), u(companyBO.stockExchangeName || '')]),
              p([chk(companyBO.listingType === 'ii'), t('  ii. It is an entity resident in jurisdictions notified by the Central Government and listed on stock exchanges in such jurisdictions, If yes, mention Stock Exchange Name: '), u(companyBO.stockExchangeName2 || '')]),
              p([chk(companyBO.listingType === 'iii'), t('  iii. It is a subsidiary of such listed entities (i & ii)')]),
              p([chk(companyBO.listingType === 'B'), t('  B. if \'A\' is not applicable, kindly share the information of the beneficial owner.')]),
            ])
          ]
        }),
      ]
    }),
    pBlank(),
    p([t('The Legal Entity as stated above hereby confirms and declares the following on the date: '), u(fmtDate(companyBO.date))]),
    pBlank(),
    p([t('Tick the box as applicable (Refer notes)', { italics: true })]),
    p([chk(companyBO.category === '1'), t('  ', { bold: true }), t('Category 1', { bold: true }), t(' - We hereby declare that following persons/ entity noted in the below table own 10% or more interest or possess the right to control management/policy decisions.')]),
    p([chk(companyBO.category === '2'), t('  ', { bold: true }), t('Category 2', { bold: true }), t(' - We hereby declare that no natural person is identified as per category 1.')]),
    pBlank(),
    p([t('The details of beneficial owner(s) is/are as follows (Table 1) (If applicable)')]),
    pBlank(),
    new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [300, 900, 1400, 700, 700, 900, 1300, 700, 1126], rows: boTableRows }),
    pBlank(),
    p([t('B. List of Senior Management Official (Mandatory)', { bold: true })]),
    pBlank(),
    new Table({ width: { size: 5000, type: WidthType.DXA }, columnWidths: [2500, 2500], rows: mgmtRows }),
    pBlank(),
    p([t('C. PEP Declaration (Mandatory)', { bold: true })]),
    pBlank(),
    p([t('Our personnel(s), partner(s), director(s), officer(s), or our family member(s) or our close associate(s) and beneficial owners, is a Politically Exposed Person (as defined by RBI)')]),
    p([chk(companyBO.pep === 'YES'), t('  YES')]),
    p([chk(companyBO.pep === 'NO' || !companyBO.pep), t('  NO')]),
    pBlank(),
    p([t('Authorised Signatory/ies:')]),
    p([u(companyBO.authSignatoryName), t('  (Name, Signature with Stamp)')]),
    pBlank(),
    p([t('This document has to be signed by the authorised signatory mentioned in the Board Resolution. Company Seal is mandatory.', { bold: true })]),
    pBlank(),
    new Paragraph({ children: [t('─'.repeat(80))], spacing: { after: 80 } }),
    pBlank(),
    p([t('#Notes:', { bold: true })]),
    p([t('a. RBI guidelines for identification of Beneficial owners', { bold: true })]),
    p([t('   Category 1: Controlling ownership interest means: >10% for Companies (Public, Private) & LLP')]),
    p([t('   Category 2: Where no natural person is identified under category 1, the beneficial owner is the relevant natural person who holds the position of senior managing official in that entity.')]),
    p([t('b. Signature on the Declaration form: person who is authorised to sign BO declaration: CS / Authorised signatory', { bold: true })]),
    p([t('c. Other Instructions', { bold: true })]),
    p([t('   1. Proof of Identity – Individual: PAN*')]),
    p([t('   2. Proof of Address – Individual: Voter ID / Driving License / Passport / Redacted Aadhaar')]),
    p([t('   3. PAN Number to be provided for Residents/Entities registered in India.')]),
  ];

  const header = makeHeader(lh);
  return new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      ...(header ? { headers: { default: header } } : {}),
      children
    }]
  });
}

// ─── Routes ─────────────────────────────────────────────────────────────────
app.post('/api/generate/mdf', async (req, res) => {
  try {
    const buf = await Packer.toBuffer(generateMDF(req.body));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="MDF_${req.body.business.name.replace(/\s+/g, '_')}.docx"`);
    res.send(buf);
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post('/api/generate/resolution', async (req, res) => {
  try {
    const buf = await Packer.toBuffer(generateResolution(req.body));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="Resolution_${req.body.business.name.replace(/\s+/g, '_')}.docx"`);
    res.send(buf);
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post('/api/generate/partnership-bo', async (req, res) => {
  try {
    const buf = await Packer.toBuffer(generatePartnershipBO(req.body));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="BO_Partnership_${req.body.business.name.replace(/\s+/g, '_')}.docx"`);
    res.send(buf);
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post('/api/generate/company-bo', async (req, res) => {
  try {
    const buf = await Packer.toBuffer(generateCompanyBO(req.body));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="BO_Company_${req.body.business.name.replace(/\s+/g, '_')}.docx"`);
    res.send(buf);
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
