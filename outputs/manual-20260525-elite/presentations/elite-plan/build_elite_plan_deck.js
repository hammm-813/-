const pptxgen = require("/Users/kawasemika/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/pptxgenjs");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Codex";
pptx.company = "SHIFT AI";
pptx.subject = "ELITEプラン全体像動画";
pptx.title = "ELITEプラン全体像動画";
pptx.lang = "ja-JP";
pptx.theme = {
  headFontFace: "Hiragino Sans",
  bodyFontFace: "Hiragino Sans",
  lang: "ja-JP",
};
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";

const W = 13.333;
const H = 7.5;
const C = {
  white: "FFFFFF",
  ink: "2B1B3A",
  muted: "6E587B",
  pink: "E8005A",
  magenta: "B00083",
  purple: "7B18D8",
  violet: "4F0B78",
  pale: "FFE6F3",
  gold: "FFE083",
  lavender: "F3D7FF",
  deep: "33004F",
};

const slides = [
  {
    no: "01",
    title: "配布エージェントは『骨組み』",
    label: "完成品ではなく、育てる前提",
    claim: "自分の業務に合わせて肉付けする力が、ELITEの本質。",
    points: ["AIを使いこなす側へ移る", "完成品よりカスタマイズ能力を育てる", "エージェントは“新入社員”として育成する"],
    type: "hero",
    notes: "時代の話から入らせてください。今、人類はAIに使われる側と使いこなす側で大きく二極化していくフェーズに入っています。テンプレを覚えて使うだけなら誰でもできる。逆に言うと、最新AIを1日1時間、半年間触り続けるだけで上位1%の使い手になれる。こんなレアタイミングは、もう来ないかもしれません。だからこそ、皆さんに渡すエージェントは完成品ではなく骨組みです。完成品を渡されると自分の業務に合わせにくく、後から剥がす二度手間も発生します。骨だけなら、自分のニーズに合わせて肉付けする力が育ちます。エージェントを手にすることは、新入社員を1人雇うのと同じです。業務ルール・顧客情報・判断基準を覚えさせ、本当に実務で使える武器に育てていきます。",
  },
  {
    no: "02",
    title: "自分専用のナレッジを積み上げる",
    label: "毎週の成果は消えない",
    claim: "12週間で、基礎から完全オリジナルへ段階的に育てる。",
    points: ["1ヶ月目：基礎を作る", "2ヶ月目：オリジナル要素を入れる", "3ヶ月目：完全オリジナルへ仕上げる"],
    type: "timeline",
    notes: "このスライドは、皆さんが毎週積み上げが消えないという事実を視覚化したものです。エリートプランは12週間で3段階の成長を設計しています。1ヶ月目は基礎を作る、2ヶ月目はオリジナル要素を入れる、3ヶ月目は完全オリジナルに仕上げる。徐々に自分の色を加えていく構造です。そして、なぜこの積み上げが資産になるのかを次のスライドで詳しく話します。",
  },
  {
    no: "03",
    title: "時間資本理論",
    label: "時間を人的資本へ変換する",
    claim: "収益化は、時間→スキル→信用→収益の順番で進む。",
    points: ["時間", "人的資本", "社会資本", "金融資本"],
    type: "flow",
    notes: "人間が最初に持っているのは時間という資本だけです。これを順番通りに変換していかないと、収益までたどり着きません。時間を人的資本、つまりスキル・知識・経験に変えます。次に、人的資本が積み上がると社会資本、つまり信用・評判に変わります。そして社会資本が積み上がって、初めて金融資本、つまり収益に転換できます。この順番は飛ばせません。エリートプランの12週間は、まさに時間から人的資本への変換作業です。",
  },
  {
    no: "04",
    title: "エージェント連携の考え方",
    label: "1体で完結させない",
    claim: "複数のAIが連携すれば、一人で組織を持つ働き方に近づく。",
    points: ["秘書", "採用", "CS", "リサーチ"],
    type: "network",
    notes: "AIは1体で完結させるだけではありません。秘書、採用、CS、リサーチなど複数のエージェントが連携することで、一人で組織を持つような働き方ができるようになります。抽象的な話で終わると腹落ちしないので、次のスライドで秘書エージェントを例に具体的なカスタマイズイメージをお見せします。",
  },
  {
    no: "05",
    title: "秘書エージェント4つのカスタマイズ例",
    label: "自分の代わりに動く状態へ",
    claim: "クセ・ルール・顧客情報・共有先を覚えさせる。",
    points: ["返信のクセを学習", "会社ロゴ入り資料", "CRM情報を反映", "BCCルールを自動化"],
    type: "quad",
    notes: "具体例で話します。秘書エージェントを4つでカスタマイズすると、自分の代わりに動いてくれます。1つ目は自分のメール返信のクセを学習させること。2つ目は会社のロゴを入れたデザインMDを用意し、社外向け資料を会社用テンプレで出力させること。3つ目はCRMを読みに行き、お客様の好みをメールに反映させること。4つ目はBCCに必ず特定アドレスを入れるルールを設定して、社内にも自動共有することです。",
  },
  {
    no: "06",
    title: "自分の目的に合わせて作る",
    label: "出口は人それぞれ",
    claim: "業界特化AI人材は、今ほぼ空席。先行者利益が大きい。",
    points: ["副業", "フリーランス", "経営者", "社内改善", "コンサル"],
    type: "paths",
    notes: "皆さんの目的は人それぞれです。副業、フリーランス、経営者、社内改善、コンサル。どれを選んでも正解です。今、業界特化のAIコンサル人材はほぼ存在していません。医療、保険、士業など、専門知識を持ったAIコンサルを必要としているのに、対応できる人がほぼゼロです。これは先行者利益が非常に大きいタイミングです。",
  },
  {
    no: "07",
    title: "マネタイズの4つの出口",
    label: "まずは自分で実現できる状態へ",
    claim: "業務改善から外販・プロダクト化まで、出口を選べる。",
    points: ["自分の業務で使う", "自社でキャリアアップ", "外販コンサル／パッケージ化", "自社プロダクト／サービス"],
    type: "quad",
    notes: "マネタイズの出口は4つあります。1つ目は自分の業務で使う。2つ目は自社でキャリアアップ。3つ目は外販コンサルまたはパッケージ化。4つ目は自社プロダクトまたはサービスです。エリートプランのミニマムゴールは、自分でやりたいことを実現できる状態です。",
  },
  {
    no: "08",
    title: "“目標設定”より先に必要なこと",
    label: "まず動いて、価値を見る",
    claim: "目標は、価値が見えてから定まる。",
    points: ["不確信", "理解", "小さく作る", "価値が見える", "目標が定まる"],
    type: "flow",
    notes: "最初はみんな不確信の状態から始まります。AIに何ができるかよく分からない。次に理解し、小さく作る。すると価値が見え、そこで初めて目標が定まります。順番が逆ではいけません。まず目標を決めてから動くのではなく、まず動いて価値を見て、それから目標を定める。だから今、明確な目標がなくても焦らないでください。",
  },
  {
    no: "09",
    title: "自信の定義",
    label: "気合いではなく、再現性",
    claim: "自信とは、行動量と結果の関係を確信している状態。",
    points: ["特定の行動", "一定量こなす", "一定の結果が出る", "だから継続できる"],
    type: "quote",
    notes: "自信とは、特定の行動を一定量こなせば、一定の結果が出ると確信している状態のことです。つまり、自信は気合いや前向きさではありません。やればやっただけ結果が出ると心の底から信じられている状態です。エリートプランは、この自信を皆さんの中に作る設計になっています。",
  },
  {
    no: "10",
    title: "マネタイズは“売り方の暗記”ではない",
    label: "買われる理由を組み立てる",
    claim: "顧客理解→課題→価値→価格→提案の連鎖を作る。",
    points: ["顧客理解", "課題", "価値", "価格", "提案"],
    type: "flow",
    notes: "マネタイズは売り方のテンプレを覚えることではありません。テンプレで人は買いません。本質は、顧客理解、課題、価値、価格、提案の5段階の連鎖です。誰が何に困っていて、それを解決すればどんな価値になり、いくらで売れて、どう提案するか。この連鎖を自分で組めるようになることが目標です。",
  },
  {
    no: "11",
    title: "単価ステップ",
    label: "一段ずつ、現実的に上げる",
    claim: "無料モニターから月100万円コンサルまで、階段で進む。",
    points: ["無料モニター", "1,000円スポット", "月5万円継続", "月100万円コンサル"],
    type: "stairs",
    notes: "最初は無料モニターからスタートします。1人でいいから使ってみたいと言ってくれる人を見つけて、価値を提供します。次に1,000円のスポット案件。そこで実績を積んだら、月5万円の継続契約。これがマネタイズの最初の山です。そして月100万円のコンサル契約。いきなり一番上を狙うのではなく、一段ずつ確実に登ります。",
  },
  {
    no: "12",
    title: "参加者に求める姿勢",
    label: "配布物を“自分の武器”に変える",
    claim: "分解する・試す・質問する・改善する。",
    points: ["分解する", "試す", "質問する", "改善する"],
    type: "quad",
    notes: "配布されたものを自分の環境でどう使うか、誰に価値があるかを常に考え、自ら手を動かし続けることが不可欠です。具体的には、分解する、試す、質問する、改善する。この4つすべてが揃って初めて、配布されたエージェントが自分の武器になっていきます。",
  },
  {
    no: "13",
    title: "曜日サイクル7日分の『なんで』",
    label: "毎週の行動に設計意図がある",
    claim: "知る→計画→作る→育てる→連携→稼ぐ→FB。",
    points: ["月：知る", "火：計画", "水：POC作成", "木：育てる", "金：連携", "土：マネタイズ", "日：発表とFB"],
    type: "week",
    notes: "月曜は知る。火曜は計画。水曜はPOC作成。木曜は育てる。金曜は連携。土曜はマネタイズ戦略。日曜は発表とフィードバックです。なんとなくこなすのではなく、各曜日の設計意図を持って回してください。特に日曜のFBの場は、集合知で見落としていたパターンを並行リサーチする時間です。",
  },
  {
    no: "14",
    title: "事前準備とサポートの考え方",
    label: "本編の価値を最大化する",
    claim: "実践会は、準備済みの人が最大限吸収する場。",
    points: ["本編：60〜70分", "サポート：必要な方のみ本編後", "最低限：月・火・水の3日間を完了"],
    type: "split",
    notes: "毎回の実践会は、導入サポートの場ではありません。本編は60分から70分。事前準備をしてきた皆さんが、最大限の価値を受け取れるように設計されています。サポート時間は本編の後、必要な方だけに実施します。本編に来る前に、月曜の知る、火曜の計画、水曜の小さく作る。この3日間の動きを終えてから臨んでください。",
  },
  {
    no: "15",
    title: "90日後に目指す状態",
    label: "人的資本を持ったスタート地点へ",
    claim: "設計・実装・価値化・改善・提案ができる状態を目指す。",
    points: ["設計できる", "作れる", "価値化できる", "改善できる", "提案できる"],
    type: "five",
    notes: "90日後のゴールは5つです。設計できる、作れる、価値化できる、改善できる、提案できる。時間資本理論で言い換えるなら、12週間で皆さんは人的資本を持った状態でスタート地点に立つということです。3ヶ月で人的資本を獲得し、半年で先行者利益の位置に立つ。これが卒業後の現実的な未来像です。",
  },
  {
    no: "16",
    title: "毎週の積み上げが資産になる",
    label: "未完成でも止まらない",
    claim: "成果物は、ポートフォリオと収益化へつながる。",
    points: ["自分専用ナレッジ", "動くデモ", "提案資料", "フィードバック", "ポートフォリオ／収益化"],
    type: "stairs",
    notes: "自分専用ナレッジ、動くデモ、提案資料、フィードバック。これらを積み上げていくと、ポートフォリオと収益化に確実に届きます。時間という資本を人的資本に変え、ポートフォリオが社会資本、信用・評判に変わり、最終的に金融資本、収益に転換できます。1週目の成果物が未完成でも止まらないでください。",
  },
  {
    no: "17",
    title: "まとめ：AIエージェントを“自分の武器”にする",
    label: "12週間で、使う側へ",
    claim: "骨組みを自分用に育て、人的資本から収益化へ進む。",
    points: ["今やる：二極化時代の先行者利益", "どうやる：12週間で人的資本化", "どこへ行く：業務改善・案件獲得・提案・副業へ"],
    type: "final",
    notes: "最後にここまでの話を全部繋ぎます。なぜ今やるのか。AIに使われる側と使いこなす側で二極化する時代だからです。どうやるのか。時間資本理論に基づき、12週間で時間を人的資本に変換します。そして出口は、業務改善、案件獲得、社内提案、コンサル、企業作業、副業です。来週の月曜から、もう一度知るのスタートです。3ヶ月後、皆さんはAIエージェントを自分の武器として持った人として、新しいスタート地点に立っています。このプランを最後まで一緒に走り抜けましょう。",
  },
];

function addBg(slide, idx) {
  slide.background = { color: "F9D3F0" };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: W, h: H,
    fill: { color: idx % 3 === 0 ? "F8B4E8" : idx % 3 === 1 ? "F5C1FA" : "F6B0D8" },
    line: { color: idx % 3 === 0 ? "F8B4E8" : idx % 3 === 1 ? "F5C1FA" : "F6B0D8" },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: W, h: H,
    fill: { color: "7B18D8", transparency: 52 },
    line: { color: "7B18D8", transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.arc, {
    x: -1.6, y: -0.8, w: 9.2, h: 5.2,
    adjustPoint: 0.28,
    line: { color: "FFFFFF", transparency: 54, width: 2.2 },
    fill: { color: "FFFFFF", transparency: 100 },
    rotate: 348,
  });
  slide.addShape(pptx.ShapeType.arc, {
    x: 3.2, y: 0.2, w: 8.8, h: 5.0,
    adjustPoint: 0.34,
    line: { color: "FFFFFF", transparency: 68, width: 1.4 },
    fill: { color: "FFFFFF", transparency: 100 },
    rotate: 12,
  });
  for (let i = 0; i < 18; i++) {
    const x = (i * 0.83 + idx * 0.31) % 11.5 + 0.2;
    const y = (i * 1.37 + idx * 0.43) % 6.2 + 0.2;
    const r = i % 4 === 0 ? 0.08 : 0.04;
    slide.addShape(pptx.ShapeType.ellipse, {
      x, y, w: r, h: r,
      fill: { color: i % 3 === 0 ? C.gold : C.white, transparency: i % 3 === 0 ? 20 : 32 },
      line: { color: i % 3 === 0 ? C.gold : C.white, transparency: 100 },
    });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 6.96, w: W, h: 0.54,
    fill: { color: C.deep, transparency: 12 },
    line: { color: C.deep, transparency: 100 },
  });
}

function addHeader(slide, s) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.55, y: 0.42, w: 4.95, h: 0.58,
    fill: { color: C.white, transparency: 4 },
    line: { color: C.white, transparency: 100 },
    radius: 0.06,
  });
  slide.addText(`ELITE プラン 特別ページ  /  ${s.no}`, {
    x: 0.72, y: 0.53, w: 4.55, h: 0.28,
    fontFace: "Hiragino Sans",
    fontSize: 16,
    bold: true,
    color: C.pink,
    breakLine: false,
    fit: "shrink",
  });
}

function addTitle(slide, s, wide = false) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.55, y: 1.25, w: wide ? 9.55 : 7.65, h: 0.9,
    fill: { color: C.pink },
    line: { color: C.pink, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: wide ? 9.25 : 7.55, y: 1.25, w: 1.9, h: 0.9,
    fill: { color: C.magenta },
    line: { color: C.magenta, transparency: 100 },
  });
  slide.addText(s.title, {
    x: 0.82, y: 1.39, w: wide ? 9.65 : 7.75, h: 0.48,
    fontFace: "Hiragino Sans",
    fontSize: s.title.length > 24 ? 24 : 29,
    bold: true,
    color: C.white,
    margin: 0,
    fit: "shrink",
  });
}

function addClaim(slide, s) {
  slide.addText(s.label, {
    x: 0.62, y: 2.46, w: 4.7, h: 0.34,
    fontFace: "Hiragino Sans",
    fontSize: 14,
    bold: true,
    color: C.pink,
    fit: "shrink",
  });
  slide.addText(s.claim, {
    x: 0.62, y: 2.84, w: 6.8, h: 1.08,
    fontFace: "Hiragino Sans",
    fontSize: 25,
    bold: true,
    color: C.ink,
    fit: "shrink",
    valign: "mid",
    breakLine: false,
  });
}

function addFooter(slide) {
  slide.addText("ELITE Plan / Agent Customization Program", {
    x: 0.58, y: 7.08, w: 4.4, h: 0.18,
    fontFace: "Avenir Next",
    fontSize: 8,
    color: "FBD7FF",
    charSpace: 1.2,
  });
}

function addAgentVisual(slide) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 9.1, y: 1.05, w: 3.75, h: 3.75,
    fill: { color: C.gold, transparency: 63 },
    line: { color: C.gold, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 9.65, y: 1.45, w: 0.68, h: 0.68,
    fill: { color: "2A1D36" },
    line: { color: "2A1D36", transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.arc, {
    x: 9.38, y: 1.22, w: 1.1, h: 0.8,
    adjustPoint: 0.35,
    line: { color: "191121", width: 6 },
    fill: { color: "191121", transparency: 100 },
    rotate: 18,
  });
  slide.addShape(pptx.ShapeType.chevron, {
    x: 9.35, y: 2.18, w: 1.45, h: 1.55,
    fill: { color: "24214C" },
    line: { color: "24214C", transparency: 100 },
    rotate: 90,
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 10.55, y: 3.2, w: 1.55, h: 0.9,
    fill: { color: "EFEAF7" },
    line: { color: "D3C4E8", width: 1 },
    rotate: 0,
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 10.35, y: 4.12, w: 2.15, h: 0.16,
    fill: { color: "BFC0D6" },
    line: { color: "BFC0D6", transparency: 100 },
  });
  for (let i = 0; i < 14; i++) {
    const x = 9.0 + ((i * 0.37) % 3.5);
    const y = 0.95 + ((i * 0.53) % 3.9);
    slide.addText("✦", {
      x, y, w: 0.18, h: 0.18,
      fontFace: "Hiragino Sans",
      fontSize: i % 2 === 0 ? 10 : 7,
      bold: true,
      color: C.gold,
      margin: 0,
      rotate: i * 7,
    });
  }
}

function chip(slide, text, x, y, w, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h: opts.h || 0.56,
    rectRadius: 0.07,
    fill: { color: opts.fill || C.white, transparency: opts.transparency ?? 0 },
    line: { color: opts.line || C.white, transparency: opts.lineTransparency ?? 100 },
  });
  slide.addText(text, {
    x: x + 0.16, y: y + 0.13, w: w - 0.32, h: 0.24,
    fontFace: "Hiragino Sans",
    fontSize: opts.size || 14,
    bold: true,
    color: opts.color || C.ink,
    align: "center",
    fit: "shrink",
    margin: 0,
  });
}

function renderBullets(slide, s) {
  s.points.slice(0, 4).forEach((p, i) => {
    const y = 4.18 + i * 0.58;
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.64, y, w: 0.16, h: 0.16,
      fill: { color: i % 2 === 0 ? C.pink : C.purple },
      line: { color: i % 2 === 0 ? C.pink : C.purple, transparency: 100 },
    });
    slide.addText(p, {
      x: 0.95, y: y - 0.08, w: 5.8, h: 0.34,
      fontFace: "Hiragino Sans",
      fontSize: 17,
      bold: true,
      color: C.ink,
      fit: "shrink",
    });
  });
}

function renderFlow(slide, s) {
  const startX = 0.82;
  const y = 4.7;
  const stepW = s.points.length === 4 ? 2.0 : 1.72;
  s.points.forEach((p, i) => {
    const x = startX + i * (stepW + 0.32);
    chip(slide, p, x, y, stepW, { fill: i === s.points.length - 1 ? C.pink : C.white, color: i === s.points.length - 1 ? C.white : C.ink, h: 0.68, size: 14 });
    if (i < s.points.length - 1) {
      slide.addShape(pptx.ShapeType.triangle, {
        x: x + stepW + 0.1, y: y + 0.21, w: 0.16, h: 0.22,
        fill: { color: C.purple },
        line: { color: C.purple, transparency: 100 },
        rotate: 90,
      });
    }
  });
}

function renderTimeline(slide, s) {
  const xs = [0.85, 3.2, 5.55];
  s.points.forEach((p, i) => {
    slide.addShape(pptx.ShapeType.rect, {
      x: xs[i], y: 4.3, w: 1.9, h: 0.18,
      fill: { color: i === 0 ? C.pink : i === 1 ? C.purple : C.gold },
      line: { color: i === 0 ? C.pink : i === 1 ? C.purple : C.gold, transparency: 100 },
    });
    slide.addText(`${i + 1}`, {
      x: xs[i] + 0.64, y: 3.88, w: 0.58, h: 0.58,
      fontFace: "Avenir Next",
      fontSize: 22,
      bold: true,
      color: C.white,
      align: "center",
      valign: "mid",
      fill: { color: i === 0 ? C.pink : i === 1 ? C.purple : C.magenta },
      margin: 0,
    });
    slide.addText(p, {
      x: xs[i] - 0.06, y: 4.64, w: 2.02, h: 0.64,
      fontFace: "Hiragino Sans",
      fontSize: 15,
      bold: true,
      color: C.ink,
      align: "center",
      fit: "shrink",
    });
  });
}

function renderQuad(slide, s) {
  const coords = [[0.72, 4.0], [3.55, 4.0], [0.72, 5.05], [3.55, 5.05]];
  s.points.slice(0, 4).forEach((p, i) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: coords[i][0], y: coords[i][1], w: 2.5, h: 0.72,
      rectRadius: 0.08,
      fill: { color: i % 2 === 0 ? C.white : C.pale, transparency: 0 },
      line: { color: i % 2 === 0 ? C.white : C.pale, transparency: 100 },
    });
    slide.addText(p, {
      x: coords[i][0] + 0.16, y: coords[i][1] + 0.2, w: 2.18, h: 0.24,
      fontFace: "Hiragino Sans",
      fontSize: 14,
      bold: true,
      color: i === 0 ? C.pink : C.ink,
      align: "center",
      fit: "shrink",
      margin: 0,
    });
  });
}

function renderNetwork(slide, s) {
  const cx = 3.55, cy = 4.78;
  chip(slide, "自分", cx - 0.54, cy - 0.24, 1.08, { fill: C.pink, color: C.white, h: 0.5, size: 15 });
  const nodes = [[1.25, 4.1], [5.35, 4.1], [1.25, 5.45], [5.35, 5.45]];
  s.points.forEach((p, i) => {
    slide.addShape(pptx.ShapeType.line, {
      x: cx, y: cy, w: nodes[i][0] - cx + 0.52, h: nodes[i][1] - cy + 0.25,
      line: { color: C.white, transparency: 15, width: 1.8 },
    });
    chip(slide, p, nodes[i][0], nodes[i][1], 1.08, { fill: C.white, h: 0.52, size: 14 });
  });
}

function renderStairs(slide, s) {
  const baseY = 5.7;
  s.points.slice(0, 5).forEach((p, i) => {
    const x = 0.68 + i * 1.28;
    const h = 0.42 + i * 0.26;
    const y = baseY - h;
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: 1.18, h,
      fill: { color: i === s.points.length - 1 ? C.pink : i % 2 === 0 ? C.white : C.lavender, transparency: i === s.points.length - 1 ? 0 : 8 },
      line: { color: C.white, transparency: 100 },
    });
    slide.addText(p, {
      x: x + 0.06, y: y + 0.11, w: 1.06, h: Math.min(h - 0.08, 0.62),
      fontFace: "Hiragino Sans",
      fontSize: 11.5,
      bold: true,
      color: i === s.points.length - 1 ? C.white : C.ink,
      align: "center",
      valign: "mid",
      fit: "shrink",
      margin: 0,
    });
  });
}

function renderWeek(slide, s) {
  const colors = [C.pink, C.magenta, C.purple, C.violet, C.ink, C.pink, C.gold];
  s.points.forEach((p, i) => {
    const x = 0.65 + (i % 4) * 1.58;
    const y = 4.0 + Math.floor(i / 4) * 0.92;
    chip(slide, p, x, y, 1.36, { fill: colors[i], color: i === 6 ? C.ink : C.white, h: 0.58, size: 12 });
  });
}

function renderFive(slide, s) {
  s.points.forEach((p, i) => {
    const angle = (-90 + i * 72) * Math.PI / 180;
    const x = 3.45 + Math.cos(angle) * 2.05;
    const y = 4.8 + Math.sin(angle) * 1.25;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x - 0.52, y: y - 0.31, w: 1.04, h: 0.62,
      fill: { color: i === 0 ? C.pink : C.white, transparency: i === 0 ? 0 : 4 },
      line: { color: C.white, transparency: 100 },
    });
    slide.addText(p, {
      x: x - 0.45, y: y - 0.1, w: 0.9, h: 0.18,
      fontFace: "Hiragino Sans",
      fontSize: 10.8,
      bold: true,
      color: i === 0 ? C.white : C.ink,
      align: "center",
      fit: "shrink",
      margin: 0,
    });
  });
}

function renderPaths(slide, s) {
  s.points.forEach((p, i) => {
    const x = 0.7 + i * 1.25;
    slide.addShape(pptx.ShapeType.chevron, {
      x, y: 4.45, w: 1.18, h: 0.74,
      fill: { color: i % 2 === 0 ? C.white : C.pale, transparency: 0 },
      line: { color: C.white, transparency: 100 },
    });
    slide.addText(p, {
      x: x + 0.1, y: 4.67, w: 0.88, h: 0.18,
      fontFace: "Hiragino Sans",
      fontSize: 11.5,
      bold: true,
      color: C.ink,
      align: "center",
      fit: "shrink",
      margin: 0,
    });
  });
}

function renderQuote(slide, s) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.7, y: 4.08, w: 6.5, h: 1.12,
    fill: { color: C.white, transparency: 0 },
    line: { color: C.white, transparency: 100 },
  });
  slide.addText("「特定の行動を一定量こなせば、一定の結果が出る」と確信している状態", {
    x: 1.0, y: 4.38, w: 5.9, h: 0.38,
    fontFace: "Hiragino Sans",
    fontSize: 19,
    bold: true,
    color: C.ink,
    align: "center",
    fit: "shrink",
    margin: 0,
  });
}

function renderSplit(slide, s) {
  chip(slide, "本編", 0.8, 4.04, 2.1, { fill: C.pink, color: C.white, h: 0.68, size: 19 });
  chip(slide, "サポート", 3.25, 4.04, 2.1, { fill: C.white, color: C.ink, h: 0.68, size: 18 });
  slide.addText("事前準備済みの人が\n最大限吸収する時間", {
    x: 0.75, y: 4.95, w: 2.25, h: 0.62,
    fontFace: "Hiragino Sans",
    fontSize: 14,
    bold: true,
    color: C.ink,
    align: "center",
    fit: "shrink",
  });
  slide.addText("必要な方のみ\n本編後に実施", {
    x: 3.22, y: 4.95, w: 2.2, h: 0.62,
    fontFace: "Hiragino Sans",
    fontSize: 14,
    bold: true,
    color: C.ink,
    align: "center",
    fit: "shrink",
  });
}

function renderContent(slide, s) {
  switch (s.type) {
    case "timeline": return renderTimeline(slide, s);
    case "flow": return renderFlow(slide, s);
    case "network": return renderNetwork(slide, s);
    case "quad": return renderQuad(slide, s);
    case "paths": return renderPaths(slide, s);
    case "quote": return renderQuote(slide, s);
    case "stairs": return renderStairs(slide, s);
    case "week": return renderWeek(slide, s);
    case "split": return renderSplit(slide, s);
    case "five": return renderFive(slide, s);
    default: return renderBullets(slide, s);
  }
}

slides.forEach((s, idx) => {
  const slide = pptx.addSlide();
  addBg(slide, idx);
  addHeader(slide, s);
  addTitle(slide, s, idx === 16 || s.title.length > 21);
  addClaim(slide, s);
  renderContent(slide, s);
  addAgentVisual(slide);
  addFooter(slide);
  slide.addNotes(`No.${s.no} ${s.title}\n\n${s.notes}`);
});

pptx.writeFile({ fileName: "/Users/kawasemika/Desktop/Claudeテスト/outputs/manual-20260525-elite/presentations/elite-plan/output/2026-05-25_ELITEプラン全体像動画_スライド.pptx" });
