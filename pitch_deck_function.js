
function createPitchDeck() {
  var pres = SlidesApp.create('The Poler Team \u2014 M&A Lead Generation');
  var firstSlide = pres.getSlides()[0];
  firstSlide.remove();

  // Color palette
  var DARK = '#0f172a';
  var ACCENT = '#2563eb';
  var GREEN = '#10b981';
  var WHITE = '#ffffff';
  var GRAY = '#94a3b8';
  var LIGHT_BG = '#1e293b';
  var GOLD = '#f59e0b';

  // ==================== SLIDE 1: TITLE ====================
  var s1 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s1.getBackground().setSolidFill(DARK);

  var title1 = s1.insertTextBox('The Poler Team', 40, 140, 640, 60);
  title1.getText().getTextStyle().setFontFamily('Inter').setFontSize(42).setBold(true).setForegroundColor(WHITE);

  var sub1 = s1.insertTextBox('M&A Lead Generation', 40, 210, 640, 40);
  sub1.getText().getTextStyle().setFontFamily('Inter').setFontSize(24).setForegroundColor(ACCENT);

  var line1 = s1.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 270, 200, 4);
  line1.getFill().setSolidFill(ACCENT);
  line1.getBorder().setTransparent();

  var desc1 = s1.insertTextBox('Connecting buyers with off-market small business owners\nwho are not going through a formal sale process.', 40, 300, 600, 60);
  desc1.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setForegroundColor(GRAY);

  var prep = s1.insertTextBox('Prepared for Olympus Cosmetic Group', 40, 390, 400, 24);
  prep.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setForegroundColor(GOLD);

  var footer1 = s1.insertTextBox('Aventura, FL  \u2022  thepolerteam.com  \u2022  Confidential', 40, 490, 400, 20);
  footer1.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(GRAY);

  // ==================== SLIDE 2: WHO WE ARE ====================
  var s2 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s2.getBackground().setSolidFill(DARK);

  var h2 = s2.insertTextBox('Who We Are', 40, 30, 640, 40);
  h2.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line2 = s2.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line2.getFill().setSolidFill(ACCENT);
  line2.getBorder().setTransparent();

  var body2 = s2.insertTextBox(
    'The Poler Team is a family-operated firm based in Aventura, Florida with over 20 years of real estate experience.\n\n' +
    'Rosa Poler founded the team in residential real estate, expanding into commercial transactions. Today, Dylan Poler leads our M&A Lead Generation division \u2014 focused exclusively on connecting acquisition-minded buyers with privately-held small business owners who are not listed on the market.\n\n' +
    'We do not represent multiple competitors in the same industry. Each buyer engagement is exclusive within their sector, ensuring confidentiality and alignment of interests.\n\n' +
    'Our approach is direct, personal outreach to owner-operators \u2014 not mass marketing. We identify, contact, and qualify small business owners, then facilitate introductory calls between owners and our buyer clients.',
    40, 95, 640, 240
  );
  body2.getText().getTextStyle().setFontFamily('Inter').setFontSize(12).setForegroundColor('#cbd5e1');
  body2.getText().getParagraphStyle().setLineSpacing(130);

  // Value props
  var vp1 = s2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 360, 200, 80);
  vp1.getFill().setSolidFill(LIGHT_BG);
  vp1.getBorder().getLineFill().setSolidFill('#334155');
  var vt1 = s2.insertTextBox('Off-Market\nDeal Sourcing', 50, 375, 180, 50);
  vt1.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);
  vt1.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var vp2 = s2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 260, 360, 200, 80);
  vp2.getFill().setSolidFill(LIGHT_BG);
  vp2.getBorder().getLineFill().setSolidFill('#334155');
  var vt2 = s2.insertTextBox('Industry\nExclusivity', 270, 375, 180, 50);
  vt2.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);
  vt2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var vp3 = s2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 480, 360, 200, 80);
  vp3.getFill().setSolidFill(LIGHT_BG);
  vp3.getBorder().getLineFill().setSolidFill('#334155');
  var vt3 = s2.insertTextBox('No Bidding\nProcess', 490, 375, 180, 50);
  vt3.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);
  vt3.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var footer2 = s2.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer2.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 3: THE PROBLEM WE SOLVE ====================
  var s3 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s3.getBackground().setSolidFill(DARK);

  var h3 = s3.insertTextBox('The Problem We Solve', 40, 30, 640, 40);
  h3.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line3 = s3.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line3.getFill().setSolidFill(ACCENT);
  line3.getBorder().setTransparent();

  // Left column - The Challenge
  var ch3 = s3.insertTextBox('The Challenge', 40, 95, 310, 24);
  ch3.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setBold(true).setForegroundColor(GOLD);

  var chBody = s3.insertTextBox(
    '\u2022 Most small business owners are not actively listing their businesses for sale\n\n' +
    '\u2022 Brokers and investment bankers create competitive bidding, driving up prices\n\n' +
    '\u2022 Off-market owners are difficult to identify and even harder to engage\n\n' +
    '\u2022 Cold outreach at scale requires systems, persistence, and a personal touch',
    40, 125, 310, 200
  );
  chBody.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor('#cbd5e1');
  chBody.getText().getParagraphStyle().setLineSpacing(120);

  // Right column - Our Solution
  var sol3 = s3.insertTextBox('Our Solution', 380, 95, 310, 24);
  sol3.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setBold(true).setForegroundColor(GREEN);

  var solBody = s3.insertTextBox(
    '\u2022 We directly contact owner-operators across the country via personalized outreach\n\n' +
    '\u2022 We achieve a 33.7% response rate from small business owners \u2014 far above industry norms\n\n' +
    '\u2022 75% of respondents agree to an introductory call\n\n' +
    '\u2022 No broker, no bidding war \u2014 just a direct conversation between buyer and seller',
    380, 125, 310, 200
  );
  solBody.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor('#cbd5e1');
  solBody.getText().getParagraphStyle().setLineSpacing(120);

  // Bottom highlight box
  var hlBox = s3.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 350, 640, 80);
  hlBox.getFill().setSolidFill('#1a2744');
  hlBox.getBorder().getLineFill().setSolidFill(ACCENT);
  var hlTxt = s3.insertTextBox('In our most recent engagement, we sourced and closed a $20M acquisition\nwhere the seller did not use an investment banker or broker \u2014\nno bidding process, direct deal.', 60, 360, 600, 60);
  hlTxt.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);
  hlTxt.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var footer3 = s3.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer3.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 4: PROVEN RESULTS ====================
  var s4 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s4.getBackground().setSolidFill(DARK);

  var h4 = s4.insertTextBox('Proven Results at Scale', 40, 30, 640, 40);
  h4.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line4 = s4.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line4.getFill().setSolidFill(ACCENT);
  line4.getBorder().setTransparent();

  var sub4 = s4.insertTextBox('Results from a single client engagement in the behavioral health sector', 40, 85, 600, 20);
  sub4.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor(GRAY);

  // Metric cards - Row 1
  var cards = [
    { val: '500+', label: 'Small Business Owners\nContacted', color: ACCENT },
    { val: '33.7%', label: 'Response Rate\nfrom Owners', color: GREEN },
    { val: '75.1%', label: 'Of Respondents\nScheduled Intro Calls', color: '#8b5cf6' },
    { val: '86.5%', label: 'Of Intro Calls Led to\nAssisted Meetings', color: GOLD }
  ];

  for (var i = 0; i < cards.length; i++) {
    var cx = 40 + i * 168;
    var card = s4.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cx, 115, 155, 95);
    card.getFill().setSolidFill(LIGHT_BG);
    card.getBorder().getLineFill().setSolidFill(cards[i].color);

    var cv = s4.insertTextBox(cards[i].val, cx + 10, 122, 135, 36);
    cv.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(cards[i].color);

    var cl = s4.insertTextBox(cards[i].label, cx + 10, 162, 135, 40);
    cl.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor(GRAY);
  }

  // Funnel visualization
  var funnelH = s4.insertTextBox('Outreach Funnel \u2014 Owner-Operators Only', 40, 225, 400, 20);
  funnelH.getText().getTextStyle().setFontFamily('Inter').setFontSize(12).setBold(true).setForegroundColor(WHITE);

  var funnelSteps = [
    { val: 526, label: 'Contacted', w: 580, color: '#334155' },
    { val: 177, label: 'Responded (33.7%)', w: 430, color: ACCENT },
    { val: 133, label: 'Intro Call Scheduled (75.1%)', w: 320, color: '#7c3aed' },
    { val: 115, label: 'Assisted Meeting (86.5%)', w: 260, color: GREEN },
    { val: 16, label: 'In Active Pipeline', w: 120, color: GOLD }
  ];

  for (var j = 0; j < funnelSteps.length; j++) {
    var fy = 252 + j * 38;
    var bar = s4.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 80, fy, funnelSteps[j].w, 28);
    bar.getFill().setSolidFill(funnelSteps[j].color);
    bar.getBorder().setTransparent();

    var fLabel = s4.insertTextBox(funnelSteps[j].val + '  ' + funnelSteps[j].label, 90, fy + 4, funnelSteps[j].w - 20, 20);
    fLabel.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setBold(true).setForegroundColor(WHITE);
  }

  // Additional stats
  var addStats = s4.insertTextBox(
    'Coverage: 35+ states  \u2022  Active pipeline deals across multiple states  \u2022  $20M closed deal (no broker, no bidding)',
    40, 460, 640, 20
  );
  addStats.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(GRAY);

  var footer4 = s4.insertTextBox('The Poler Team  \u2022  Confidential  \u2022  Sanitized data from a single client engagement', 40, 500, 500, 16);
  footer4.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 5: HOW WE WORK ====================
  var s5 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s5.getBackground().setSolidFill(DARK);

  var h5 = s5.insertTextBox('How We Work', 40, 30, 640, 40);
  h5.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line5 = s5.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line5.getFill().setSolidFill(ACCENT);
  line5.getBorder().setTransparent();

  var steps = [
    { num: '01', title: 'Research & Targeting', desc: 'We identify owner-operated businesses that fit your acquisition criteria \u2014 by geography, size, specialty, and ownership type.' },
    { num: '02', title: 'Personalized Outreach', desc: 'Our team contacts owners directly via LinkedIn, phone, email, text, and voicemail. Every message is personalized. No mass blasts.' },
    { num: '03', title: 'Qualify & Schedule', desc: 'When an owner expresses interest, we qualify the opportunity and schedule an introductory call between you and the owner. No commitment required from either party.' },
    { num: '04', title: 'Facilitate & Support', desc: 'We assist through the meeting process and remain a resource as the relationship develops \u2014 from first call through LOI and beyond.' }
  ];

  for (var k = 0; k < steps.length; k++) {
    var sy = 95 + k * 100;
    var numBox = s5.insertShape(SlidesApp.ShapeType.ELLIPSE, 40, sy + 5, 40, 40);
    numBox.getFill().setSolidFill(ACCENT);
    numBox.getBorder().setTransparent();
    var numTxt = s5.insertTextBox(steps[k].num, 40, sy + 12, 40, 24);
    numTxt.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setBold(true).setForegroundColor(WHITE);
    numTxt.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var stTitle = s5.insertTextBox(steps[k].title, 95, sy, 585, 24);
    stTitle.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setBold(true).setForegroundColor(WHITE);

    var stDesc = s5.insertTextBox(steps[k].desc, 95, sy + 26, 585, 50);
    stDesc.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor('#cbd5e1');
    stDesc.getText().getParagraphStyle().setLineSpacing(120);
  }

  var footer5 = s5.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer5.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 6: WHY OFF-MARKET ====================
  var s6 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s6.getBackground().setSolidFill(DARK);

  var h6 = s6.insertTextBox('Why Off-Market Deals', 40, 30, 640, 40);
  h6.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line6 = s6.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line6.getFill().setSolidFill(ACCENT);
  line6.getBorder().setTransparent();

  // Comparison table header
  var thBg = s6.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 100, 640, 30);
  thBg.getFill().setSolidFill('#1e293b');
  thBg.getBorder().setTransparent();

  var th1 = s6.insertTextBox('', 40, 103, 200, 24);
  var th2 = s6.insertTextBox('Broker / Banker Process', 240, 103, 210, 24);
  th2.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setBold(true).setForegroundColor('#f87171');
  th2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  var th3 = s6.insertTextBox('The Poler Team', 470, 103, 210, 24);
  th3.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setBold(true).setForegroundColor(GREEN);
  th3.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var rows = [
    ['Competition', 'Multiple bidders', 'You are the only buyer'],
    ['Pricing', 'Inflated by bidding', 'Negotiate directly with owner'],
    ['Timeline', 'Lengthy formal process', 'Move at your pace'],
    ['Relationship', 'Filtered through intermediary', 'Direct owner relationship'],
    ['Deal Flow', 'Wait for listings', 'Proactive \u2014 we find them'],
    ['Confidentiality', 'Widely marketed', 'Private, targeted outreach']
  ];

  for (var r = 0; r < rows.length; r++) {
    var ry = 135 + r * 40;
    if (r % 2 === 0) {
      var rowBg = s6.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, ry, 640, 38);
      rowBg.getFill().setSolidFill('#0f172a');
      rowBg.getBorder().setTransparent();
    }
    var rc1 = s6.insertTextBox(rows[r][0], 50, ry + 8, 180, 22);
    rc1.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setBold(true).setForegroundColor(WHITE);
    var rc2 = s6.insertTextBox(rows[r][1], 245, ry + 8, 200, 22);
    rc2.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor('#94a3b8');
    rc2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    var rc3 = s6.insertTextBox(rows[r][2], 475, ry + 8, 200, 22);
    rc3.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(GREEN);
    rc3.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }

  var footer6 = s6.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer6.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 7: NEXT STEPS ====================
  var s7 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s7.getBackground().setSolidFill(DARK);

  var h7 = s7.insertTextBox('Next Steps', 40, 30, 640, 40);
  h7.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line7 = s7.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line7.getFill().setSolidFill(ACCENT);
  line7.getBorder().setTransparent();

  var nextBody = s7.insertTextBox(
    'We would welcome the opportunity to support Olympus Cosmetic Group\'s acquisition strategy.\n\n' +
    'Our proposed next steps:',
    40, 95, 640, 60
  );
  nextBody.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setForegroundColor('#cbd5e1');

  var nextSteps = [
    { icon: '\u260E', title: 'Introductory Call', desc: 'A brief call to learn about your target criteria \u2014 geography, practice type, size, and deal structure preferences.' },
    { icon: '\uD83C\uDFAF', title: 'Target List Development', desc: 'We build a curated list of owner-operated practices matching your criteria across your target states.' },
    { icon: '\uD83D\uDCE8', title: 'Outreach Campaign Launch', desc: 'Personalized, multi-channel outreach begins. You receive qualified introductions \u2014 no upfront cost until engagement.' },
    { icon: '\uD83E\uDD1D', title: 'Buyer Agreement', desc: 'Once you see the quality of our pipeline, we formalize a buyer\u2019s agreement for ongoing deal sourcing.' }
  ];

  for (var n = 0; n < nextSteps.length; n++) {
    var ny = 170 + n * 75;
    var nBox = s7.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, ny, 640, 65);
    nBox.getFill().setSolidFill(LIGHT_BG);
    nBox.getBorder().getLineFill().setSolidFill('#334155');

    var nTitle = s7.insertTextBox(nextSteps[n].title, 60, ny + 8, 580, 22);
    nTitle.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);

    var nDesc = s7.insertTextBox(nextSteps[n].desc, 60, ny + 30, 600, 30);
    nDesc.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(GRAY);
  }

  var footer7 = s7.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer7.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 8: CONTACT ====================
  var s8 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s8.getBackground().setSolidFill(DARK);

  var cTitle = s8.insertTextBox('Let\u2019s Connect', 40, 150, 640, 50);
  cTitle.getText().getTextStyle().setFontFamily('Inter').setFontSize(36).setBold(true).setForegroundColor(WHITE);

  var cLine = s8.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 210, 200, 4);
  cLine.getFill().setSolidFill(ACCENT);
  cLine.getBorder().setTransparent();

  var cName = s8.insertTextBox('Dylan Poler', 40, 240, 400, 30);
  cName.getText().getTextStyle().setFontFamily('Inter').setFontSize(18).setBold(true).setForegroundColor(WHITE);

  var cRole = s8.insertTextBox('M&A Lead Generation', 40, 270, 400, 24);
  cRole.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setForegroundColor(ACCENT);

  var cTeam = s8.insertTextBox('The Poler Team', 40, 300, 400, 24);
  cTeam.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setForegroundColor(GRAY);

  var cLoc = s8.insertTextBox('Aventura, FL', 40, 330, 400, 24);
  cLoc.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setForegroundColor(GRAY);

  var footer8 = s8.insertTextBox('Confidential  \u2022  For Olympus Cosmetic Group review only', 40, 490, 400, 20);
  footer8.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor('#475569');

  return pres.getUrl();
}
