<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ボックスカルバート水理計算書</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        window.MathJax = {
            tex: { 
                inlineMath: [['$', '$']], 
                displayMath: [['$$', '$$']],
                processEscapes: true
            },
            chtml: { scale: 0.95 },
            svg: { fontCache: 'global' }
        };
    </script>
    <script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;600&family=Noto+Sans+JP:wght@400;700&display=swap');
        
        body { font-family: 'Noto Sans JP', sans-serif; background-color: #e2e8f0; margin: 0; }
        .font-sans { font-family: 'Noto Sans JP', sans-serif; }
        .font-serif { font-family: 'Noto Serif JP', serif; }

/* =================================================================
   [NEW] 修正後
   ================================================================= */
/* =================================================================
   [NEW] 今回の修正版（全体を上書きしてください）
   ================================================================= */
        /* 印刷設定 */
        @media print {
            @page { 
                size: A4; 
                margin: 0; 
            }

            .no-print { display: none !important; }

            /* html, body, main すべての高さを解放 */
            html, body, main {
                background: white;
                display: block !important;
                height: auto !important;
                width: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: visible !important;
            }

            .report-page { 
                margin: 0 auto !important;
                padding: 10mm !important; /* コンテンツの余白 */
                width: 100% !important;
                
                /* 高さ制限を解除 */
                min-height: 0 !important; 
                height: auto !important;
                
                border: none !important;
                box-shadow: none !important;
                
                /* [重要] 強制改ページを削除 */
                page-break-after: auto !important; 
                page-break-inside: avoid !important; /* ページ途中で切れないようにする */
                
                transform: none !important;
            }

            /* 色調整 */
            .bg-slate-50 { background-color: #f8fafc !important; -webkit-print-color-adjust: exact; }
            .bg-blue-50 { background-color: #eff6ff !important; -webkit-print-color-adjust: exact; }
            .border-slate-300 { border-color: #cbd5e1 !important; -webkit-print-color-adjust: exact; }
        }

        /* プレビュー画面（A4縦） */
        .report-page {
            background: white;
            width: 210mm;
            min-height: 297mm;
            padding: 10mm 15mm; 
            margin: 20px auto;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            box-sizing: border-box;
        }

        /* 数式ブロック */
        .math-block { 
            @apply bg-slate-50 px-2 py-1 rounded border border-slate-200 my-0.5 text-[9px] leading-tight overflow-x-auto font-serif;
        }
        
        /* 公式用の数式ブロック（大きめ） */
        .math-block-lg {
            @apply bg-slate-50 px-3 py-1.5 rounded border border-slate-200 my-0.5 text-[11px] leading-snug overflow-x-auto font-serif font-bold;
        }

        /* 入力フィールド */
        .input-field { 
            @apply border border-slate-300 rounded p-1.5 outline-none transition text-sm font-bold font-sans;
            background-color: #ffffff !important;
            color: #0f172a !important;
        }
        .input-field:focus { @apply ring-2 ring-blue-500 border-blue-500; }
        
        /* 幅指定 */
        input[type="number"].input-field { width: 50% !important; }
        select.input-field { width: 100% !important; }
        
        .label-text { @apply block text-[10px] font-bold text-slate-300 mb-1 uppercase tracking-wider font-sans; }
        
        /* ボタン */
        .ratio-btn { @apply flex-1 py-1 text-[10px] bg-slate-700 text-slate-300 border border-slate-600 hover:bg-slate-600 transition font-sans; }

        /* テーブル微調整 */
        td { padding-top: 1px; padding-bottom: 1px; }
    </style>
</head>
<body class="text-slate-900 h-screen overflow-hidden flex flex-col lg:flex-row">

    <!-- 左サイドバー：入力エリア -->
    <aside class="no-print w-full lg:w-72 bg-slate-900 text-white p-5 overflow-y-auto shrink-0 z-20 shadow-2xl">
        <h2 class="text-base font-bold border-b border-slate-700 pb-2 mb-4 flex items-center gap-2 font-sans">
            <span class="w-1 h-4 bg-blue-500 rounded"></span>
            水理条件設定
        </h2>

        <div class="space-y-4">
            <div>
                <label class="label-text">流速公式</label>
                <select id="formula-type" class="input-field" onchange="calculate()">
                    <option value="kutter" selected>クッター公式 (Kutter)</option>
                    <option value="manning">マニング公式 (Manning)</option>
                </select>
            </div>

            <div class="grid grid-cols-2 gap-2">
                <div>
                    <label class="label-text">内幅 B (m)</label>
                    <input type="number" id="B" class="input-field" value="2.000" step="0.1" oninput="calculate()">
                </div>
                <div>
                    <label class="label-text">内高 H (m)</label>
                    <input type="number" id="H" class="input-field" value="2.000" step="0.1" oninput="calculate()">
                </div>
                <div>
                    <label class="label-text">下ハンチ c1 (m)</label>
                    <input type="number" id="c1" class="input-field" value="0.200" step="0.05" oninput="calculate()">
                </div>
                <div>
                    <label class="label-text">上ハンチ c2 (m)</label>
                    <input type="number" id="c2" class="input-field" value="0.200" step="0.05" oninput="calculate()">
                </div>
            </div>

            <div class="p-2 bg-slate-800 rounded border border-slate-700">
                <label class="label-text text-blue-300">検討水深 h (m)</label>
                <input type="number" id="h" class="input-field text-lg !text-blue-700 mb-2" value="1.600" step="0.001" oninput="calculate()">
                <div class="flex">
                    <button onclick="setRatio(0.8)" class="ratio-btn rounded-l">80%</button>
                    <button onclick="setRatio(0.9)" class="ratio-btn">90%</button>
                    <button onclick="setRatio(1.0)" class="ratio-btn rounded-r">満水</button>
                </div>
            </div>

            <div class="grid grid-cols-2 gap-2">
                <div>
                    <label class="label-text">粗度係数 n</label>
                    <input type="number" id="n" class="input-field" value="0.013" step="0.001" oninput="calculate()">
                </div>
                <div>
                    <label class="label-text">勾配 I (%)</label>
                    <input type="number" id="I_pct" class="input-field" value="1.0" step="0.1" oninput="calculate()">
                </div>
            </div>

            <button onclick="window.print()" class="w-full mt-6 bg-blue-600 hover:bg-blue-500 text-white font-bold py-2 rounded shadow transition flex items-center justify-center gap-2 font-sans text-xs">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2z"></path></svg>
                PDF出力 / 印刷
            </button>
        </div>
    </aside>

    <!-- メイン：計算書プレビュー -->
    <main class="flex-1 overflow-y-auto p-4 md:p-8 bg-slate-200">
        <div class="report-page font-serif flex flex-col text-slate-900 text-[10px]">
            
            <div class="text-center border-b border-slate-900 pb-1 mb-2">
                <h1 class="text-lg font-bold tracking-widest">水理計算書</h1>
                <p class="text-[9px] font-sans text-slate-500">（ボックスカルバート）</p>
            </div>

            <!-- 1. 設計条件 -->
            <section class="mb-2">
                <h2 class="text-[10px] font-bold bg-slate-800 text-white px-2 py-0.5 mb-1 inline-block print:text-black print:bg-transparent print:border-b print:border-black print:px-0 font-sans">1. 設計条件</h2>
                
                <div class="flex flex-row gap-3 items-start">
                    <div class="w-1/2">
                        <table class="w-full text-[9px] border-collapse font-sans leading-tight">
                            <tr class="border-b border-slate-200"><td class="py-0.5 text-slate-500">採用公式</td><td class="text-right font-bold text-black"><span id="out-formula"></span></td></tr>
                            <tr class="border-b border-slate-200"><td class="py-0.5 text-slate-500">内幅 $B$</td><td class="text-right font-bold"><span id="out-B"></span> m</td></tr>
                            <tr class="border-b border-slate-200"><td class="py-0.5 text-slate-500">内高 $H$</td><td class="text-right font-bold"><span id="out-H"></span> m</td></tr>
                            <tr class="border-b border-slate-200"><td class="py-0.5 text-slate-500">下ハンチ $c_1$</td><td class="text-right font-bold"><span id="out-c1"></span> m</td></tr>
                            <tr class="border-b border-slate-200"><td class="py-0.5 text-slate-500">上ハンチ $c_2$</td><td class="text-right font-bold"><span id="out-c2"></span> m</td></tr>
                            <tr class="border-b border-slate-200 bg-blue-50/50"><td class="py-0.5 text-blue-800 font-bold">検討水深 $h$</td><td class="text-right font-bold text-blue-800"><span id="out-h"></span> m</td></tr>
                            <tr class="border-b border-slate-200"><td class="py-0.5 text-slate-500">粗度係数 $n$</td><td class="text-right font-bold"><span id="out-n"></span></td></tr>
                            <tr class="border-b border-slate-200"><td class="py-0.5 text-slate-500">水路勾配 $I$</td><td class="text-right font-bold"><span id="out-I"></span> %</td></tr>
                        </table>
                    </div>

                    <div class="w-1/2 flex flex-col items-center">
                        <div id="svg-container" class="bg-white p-1 rounded border border-slate-100 mb-0.5 w-full flex justify-center h-[140px]"></div>
                        <p class="text-[8px] text-slate-500 font-sans mt-1">図-1 計算断面図</p>
                    </div>
                </div>
            </section>

            <!-- 2. 断面定数 -->
            <section class="mb-2">
                <h2 class="text-[10px] font-bold bg-slate-800 text-white px-2 py-0.5 mb-1 inline-block print:text-black print:bg-transparent print:border-b print:border-black print:px-0 font-sans">2. 断面定数の算出</h2>
                
                <div class="ml-2 space-y-1 font-sans">
                    <!-- 断面積 -->
                    <div>
                        <div class="flex justify-between items-end mb-0.5">
                            <h3 class="text-[10px] font-bold text-slate-800 border-l-2 border-slate-300 pl-1">2.1 通水断面積 $A$ ($m^2$)</h3>
                            <span class="text-[9px] text-slate-500">公式: $A = B \cdot h - \sum \Delta A$</span>
                        </div>
                        <div class="text-[9px] text-slate-600 ml-2 font-mono leading-none tracking-tight">
                            <span id="desc-dA1"></span>, <span id="desc-dA2"></span>
                        </div>
                        <div id="math-A" class="math-block"></div>
                    </div>

                    <!-- 潤辺 -->
                    <div>
                        <div class="flex justify-between items-end mb-0.5">
                            <h3 class="text-[10px] font-bold text-slate-800 border-l-2 border-slate-300 pl-1">2.2 潤辺 $P$ ($m$)</h3>
                            <span class="text-[9px] text-slate-500">公式: $P = (B + 2h) + \sum \Delta P$</span>
                        </div>
                        <div class="text-[9px] text-slate-600 ml-2 font-mono leading-none tracking-tight">
                            <span id="desc-dP1"></span>, <span id="desc-dP2"></span>
                        </div>
                        <div id="math-P" class="math-block"></div>
                    </div>

                    <!-- 径深 -->
                    <div>
                        <h3 class="text-[10px] font-bold text-slate-800 border-l-2 border-slate-300 pl-1 mb-0.5">2.3 径深 $R$ ($m$)</h3>
                        <span class="text-[9px] text-slate-500 ml-2">公式: $R = A / P$</span>
                        <div id="math-R" class="math-block py-1"></div>
                    </div>
                </div>
            </section>

            <!-- 3. 流速・流量 -->
            <!-- flex-grow削除 -->
            <section class="mb-1">
                <h2 class="text-[10px] font-bold bg-slate-800 text-white px-2 py-0.5 mb-1 inline-block print:text-black print:bg-transparent print:border-b print:border-black print:px-0 font-sans">3. 平均流速および流量</h2>
                
                <div id="velocity-logic" class="ml-2 font-sans">
                    <!-- JSで数式挿入 -->
                </div>
            </section>

            <!-- 4. 結果 (余白なしで直結) -->
            <section class="mt-0 pt-0 border-t border-slate-400 mb-0">
                <div class="flex justify-between items-center mb-0.5 pt-1">
                    <h2 class="text-[10px] font-bold text-slate-800 font-sans">4. 計算結果</h2>
                    <span class="text-[9px] text-slate-500 font-sans">公式: $Q = A \cdot V$</span>
                </div>
                <div id="math-Q" class="bg-blue-50 border border-blue-200 p-1.5 text-center rounded font-bold text-blue-900 text-xs font-serif leading-none"></div>
            </section>

        </div>
    </main>

    <script>
        function setRatio(r) {
            const H = parseFloat(document.getElementById('H').value);
            document.getElementById('h').value = (H * r).toFixed(3);
            calculate();
        }

        async function calculate() {
            const formula = document.getElementById('formula-type').value;
            const B = parseFloat(document.getElementById('B').value) || 0;
            const H = parseFloat(document.getElementById('H').value) || 0;
            const c1 = parseFloat(document.getElementById('c1').value) || 0;
            const c2 = parseFloat(document.getElementById('c2').value) || 0;
            const h = parseFloat(document.getElementById('h').value) || 0;
            const n = parseFloat(document.getElementById('n').value) || 0.013;
            const Ip = parseFloat(document.getElementById('I_pct').value) || 0;
            const I = Ip / 100;

            // UI更新
            document.getElementById('out-formula').innerText = (formula === 'manning') ? 'マニング公式' : 'クッター公式';
            document.getElementById('out-B').innerText = B.toFixed(3);
            document.getElementById('out-H').innerText = H.toFixed(3);
            document.getElementById('out-c1').innerText = c1.toFixed(3);
            document.getElementById('out-c2').innerText = c2.toFixed(3);
            document.getElementById('out-h').innerText = h.toFixed(3);
            document.getElementById('out-n').innerText = n.toFixed(3);
            document.getElementById('out-I').innerText = Ip.toFixed(2);

            // --- 1. 断面計算 A ---
            let dA1_val = 0, dA2_val = 0;
            let dA1_desc = "", dA2_desc = "";

            if (h >= c1) {
                dA1_val = c1 * c1; 
                dA1_desc = `$\\Delta A_1 = c_1^2 = ${c1.toFixed(3)}^2 = ${dA1_val.toFixed(3)}$`;
            } else {
                dA1_val = h * h;
                dA1_desc = `$\\Delta A_1 = h^2 = ${h.toFixed(3)}^2 = ${dA1_val.toFixed(3)}$ (部分)`;
            }

            if (h > (H - c2)) {
                let th = h - (H - c2);
                dA2_val = th * th;
                dA2_desc = `$\\Delta A_2 = \\Delta h^2 = ${th.toFixed(3)}^2 = ${dA2_val.toFixed(3)}$ (部分)`;
            } else {
                dA2_val = 0;
                dA2_desc = `$\\Delta A_2 = 0$`;
            }
            const A = (B * h) - dA1_val - dA2_val;
            
            document.getElementById('desc-dA1').innerHTML = dA1_desc;
            document.getElementById('desc-dA2').innerHTML = dA2_desc;

            // --- 2. 潤辺 P ---
            let dP1_val = 0, dP2_val = 0;
            let dP1_desc = "", dP2_desc = "";

            if (h >= c1) {
                dP1_val = 2 * (Math.sqrt(2) * c1 - 2 * c1);
                dP1_desc = `$\\Delta P_1 = 2c_1(\\sqrt{2}-2) = ${dP1_val.toFixed(3)}$`;
            } else {
                dP1_val = 2 * (Math.sqrt(2) * h - 2 * h);
                dP1_desc = `$\\Delta P_1 = 2h(\\sqrt{2}-2) = ${dP1_val.toFixed(3)}$`;
            }

            if (h > (H - c2)) {
                let th = h - (H - c2);
                dP2_val = 2 * (Math.sqrt(2) * th - th); 
                dP2_desc = `$\\Delta P_2 = 2\\Delta h(\\sqrt{2}-1) = ${dP2_val.toFixed(3)}$`;
            } else {
                dP2_val = 0;
                dP2_desc = `$\\Delta P_2 = 0$`;
            }

            let P = B + 2 * h + dP1_val + dP2_val;
            
            let top_add = 0;
            if (h >= H) {
                top_add = B - 2 * c2;
                P += top_add;
                dP2_desc += `, 天井 +${top_add.toFixed(3)}`;
            }

            document.getElementById('desc-dP1').innerHTML = dP1_desc;
            document.getElementById('desc-dP2').innerHTML = dP2_desc;

            const R = (P > 0) ? A / P : 0;

            document.getElementById('math-A').innerHTML = `
            $$ \\begin{aligned} A &= ${B.toFixed(3)} \\cdot ${h.toFixed(3)} - (${dA1_val.toFixed(4)} + ${dA2_val.toFixed(4)}) \\\\
            &= ${A.toFixed(4)} \\text{ m}^2 \\end{aligned} $$`;

            document.getElementById('math-P').innerHTML = `
            $$ \\begin{aligned} P &= (${B.toFixed(3)} + 2 \\cdot ${h.toFixed(3)}) ${dP1_val < 0 ? '-' : '+'} ${Math.abs(dP1_val).toFixed(4)} + ${dP2_val.toFixed(4)} ${h>=H ? '+ ' + top_add.toFixed(3) : ''} \\\\
            &= ${P.toFixed(4)} \\text{ m} \\end{aligned} $$`;

            document.getElementById('math-R').innerHTML = `
            $$ R = \\frac{A}{P} = \\frac{${A.toFixed(4)}}{${P.toFixed(4)}} = ${R.toFixed(4)} \\text{ m} $$`;

            // --- 流速・流量 ---
            let V = 0;
            let logicHTML = "";

            if (formula === 'kutter') {
                const num = 23 + (1/n) + (0.00155/I);
                const den_part = (23 + (0.00155/I)) * (n / Math.sqrt(R));
                const den = 1 + den_part;
                const C = num / den;
                V = C * Math.sqrt(R * I);

                logicHTML = `
                <div class="mb-1">
                    <div class="flex justify-between items-end mb-0.5">
                        <h3 class="text-[10px] font-bold text-slate-800 mb-0.5 border-l-2 border-slate-300 pl-1">3.1 流速係数 $C$ (Kutter)</h3>
                        <div class="math-block-lg">
                            $$C = \\frac{23 + \\frac{1}{n} + \\frac{0.00155}{I}}{1 + \\left( 23 + \\frac{0.00155}{I} \\right) \\frac{n}{\\sqrt{R}}}$$
                        </div>
                    </div>
                    <div class="math-block">
                    $$ \\begin{aligned} 
                    C &= \\frac{23 + \\frac{1}{${n.toFixed(3)}} + \\frac{0.00155}{${I.toFixed(5)}}}{1 + \\left( 23 + \\frac{0.00155}{${I.toFixed(5)}} \\right) \\frac{${n.toFixed(3)}}{\\sqrt{${R.toFixed(4)}}}} \\\\
                    &= ${C.toFixed(4)}
                    \\end{aligned} $$
                    </div>
                </div>
                <div>
                    <div class="flex justify-between items-end mb-0.5">
                        <h3 class="text-[10px] font-bold text-slate-800 mb-0.5 border-l-2 border-slate-300 pl-1">3.2 平均流速 $V$</h3>
                        <div class="math-block-lg">
                            $$V = C \\sqrt{R \\cdot I}$$
                        </div>
                    </div>
                    <div class="math-block">
                    $$ \\begin{aligned}
                    V &= ${C.toFixed(4)} \\cdot \\sqrt{${R.toFixed(4)} \\cdot ${I.toFixed(5)}} \\\\
                    &= ${V.toFixed(4)} \\text{ m/s}
                    \\end{aligned} $$
                    </div>
                </div>`;

            } else {
                V = (1/n) * Math.pow(R, 2/3) * Math.pow(I, 1/2);
                logicHTML = `
                <div>
                    <div class="flex justify-between items-end mb-0.5">
                        <h3 class="text-[10px] font-bold text-slate-800 mb-0.5 border-l-2 border-slate-300 pl-1">3.1 平均流速 $V$ (Manning)</h3>
                        <div class="math-block-lg">
                            $$V = \\frac{1}{n} R^{2/3} I^{1/2}$$
                        </div>
                    </div>
                    <div class="math-block">
                    $$ \\begin{aligned}
                    V &= \\frac{1}{${n.toFixed(3)}} \\cdot ${R.toFixed(4)}^{2/3} \\cdot ${I.toFixed(5)}^{1/2} \\\\
                    &= ${V.toFixed(4)} \\text{ m/s}
                    \\end{aligned} $$
                    </div>
                </div>`;
            }

            document.getElementById('velocity-logic').innerHTML = logicHTML;

            const Q = A * V;
            document.getElementById('math-Q').innerHTML = `
            $$ Q = A \\cdot V = ${A.toFixed(4)} \\cdot ${V.toFixed(4)} = ${Q.toFixed(4)} \\text{ m}^3/\\text{s} $$`;

            renderSVG(B, H, h, c1, c2);
            if (window.MathJax && MathJax.typesetPromise) await MathJax.typesetPromise();
        }

        function renderSVG(B, H, h, c1, c2) {
            const container = document.getElementById('svg-container');
            const svgW = 220, svgH = 140, pad = 20; 
            const scale = Math.min((svgW - pad*2)/B, (svgH - pad*2)/H);
            
            const w = B * scale;
            const ht = H * scale;
            const x0 = (svgW - w)/2 + 5;
            const y0 = (svgH - ht)/2;
            
            const s1 = c1 * scale;
            const s2 = c2 * scale;
            const wh = Math.min(h, H) * scale;
            const wy = y0 + ht - wh;

            const box = `M ${x0} ${y0+s2} L ${x0+s2} ${y0} L ${x0+w-s2} ${y0} L ${x0+w} ${y0+s2} L ${x0+w} ${y0+ht-s1} L ${x0+w-s1} ${y0+ht} L ${x0+s1} ${y0+ht} L ${x0} ${y0+ht-s1} Z`;
            
            let poly = [];
            poly.push([x0+s1, y0+ht]);
            poly.push([x0+w-s1, y0+ht]);
            
            if (h <= c1) {
                let dx = h * scale;
                poly.push([x0+w-s1+dx, wy]);
                poly.push([x0+s1-dx, wy]);
            } else {
                poly.push([x0+w, y0+ht-s1]); 
                if (h > (H - c2)) {
                    poly.push([x0+w, y0+s2]); 
                    let th = (h - (H - c2)) * scale;
                    poly.push([x0+w-th, wy]); 
                    poly.push([x0+th, wy]);   
                    poly.push([x0, y0+s2]);   
                } else {
                    poly.push([x0+w, wy]);
                    poly.push([x0, wy]);
                }
                poly.push([x0, y0+ht-s1]);
            }
            const waterPath = "M " + poly.map(p => p[0]+","+p[1]).join(" L ") + " Z";

            const dimColor = "#64748b";
            const dimTxt = "#475569";
            
            const lineB = `
                <line x1="${x0}" y1="${y0+ht+10}" x2="${x0+w}" y2="${y0+ht+10}" stroke="${dimColor}" stroke-width="0.8"/>
                <line x1="${x0}" y1="${y0+ht+5}" x2="${x0}" y2="${y0+ht+13}" stroke="${dimColor}" stroke-width="0.5"/>
                <line x1="${x0+w}" y1="${y0+ht+5}" x2="${x0+w}" y2="${y0+ht+13}" stroke="${dimColor}" stroke-width="0.5"/>
                <text x="${x0+w/2}" y="${y0+ht+19}" text-anchor="middle" font-size="9" fill="${dimTxt}">B=${B.toFixed(2)}</text>
            `;
            const lineH = `
                <line x1="${x0-10}" y1="${y0}" x2="${x0-10}" y2="${y0+ht}" stroke="${dimColor}" stroke-width="0.8"/>
                <line x1="${x0-5}" y1="${y0}" x2="${x0-13}" y2="${y0}" stroke="${dimColor}" stroke-width="0.5"/>
                <line x1="${x0-5}" y1="${y0+ht}" x2="${x0-13}" y2="${y0+ht}" stroke="${dimColor}" stroke-width="0.5"/>
                <text x="${x0-15}" y="${y0+ht/2}" text-anchor="middle" transform="rotate(-90 ${x0-15} ${y0+ht/2})" font-size="9" fill="${dimTxt}">H=${H.toFixed(2)}</text>
            `;
            
            // h寸法: 中央水面
            const line_h = `
                <line x1="${x0}" y1="${wy}" x2="${x0+w}" y2="${wy}" stroke="#2563eb" stroke-width="1" stroke-dasharray="3,2"/>
                <text x="${x0+w/2}" y="${wy-3}" text-anchor="middle" font-size="10" font-weight="bold" fill="#2563eb">h=${h.toFixed(3)}</text>
            `;
            
            // ハンチ
            const line_c1 = `
                <line x1="${x0+s1}" y1="${y0+ht}" x2="${x0-20}" y2="${y0+ht+8}" stroke="#94a3b8" stroke-width="0.5"/>
                <text x="${x0-22}" y="${y0+ht+10}" font-size="8" fill="${dimTxt}" text-anchor="end">c1</text>
            `;
            const line_c2 = `
                <line x1="${x0+s2}" y1="${y0}" x2="${x0-20}" y2="${y0-8}" stroke="#94a3b8" stroke-width="0.5"/>
                <text x="${x0-22}" y="${y0-6}" font-size="8" fill="${dimTxt}" text-anchor="end">c2</text>
            `;

            container.innerHTML = `
                <svg width="${svgW}" height="${svgH}" viewBox="0 0 ${svgW} ${svgH}" style="font-family:sans-serif">
                    <path d="${waterPath}" fill="#bfdbfe" />
                    <path d="${box}" fill="none" stroke="#1e293b" stroke-width="1.5" />
                    ${lineB} ${lineH} ${line_h} ${line_c1} ${line_c2}
                </svg>`;
        }

        window.onload = calculate;
    </script>
</body>
</html>