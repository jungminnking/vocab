"""
Regenerates index.html from Key_Expressions.xlsx.
Run automatically by .github/workflows/build.yml whenever the Excel file changes.
Can also be run manually:  python build.py
"""
import pandas as pd, json, re
from collections import Counter
from itertools import product

EXCEL_PATH = "Key_Expressions.xlsx"
SHEET_NAME = "Vocab"
TEMPLATE_PATH = "template.html"
OUTPUT_PATH = "index.html"

# ---- irregular verb map: base -> inflected forms ----
IRREGULAR = {
 'be':['am','is','are','was','were','been','being'],
 'become':['became','become','becoming'],
 'begin':['began','begun','beginning'],
 'break':['broke','broken','breaking'],
 'bring':['brought','bringing'],
 'build':['built','building'],
 'buy':['bought','buying'],
 'catch':['caught','catching'],
 'choose':['chose','chosen','choosing'],
 'come':['came','coming'],
 'cut':['cutting'],
 'deal':['dealt','dealing'],
 'dig':['dug','digging'],
 'do':['does','did','done','doing'],
 'draw':['drew','drawn','drawing'],
 'drink':['drank','drunk','drinking'],
 'drive':['drove','driven','driving'],
 'eat':['ate','eaten','eating'],
 'fall':['fell','fallen','falling'],
 'feed':['fed','feeding'],
 'feel':['felt','feeling'],
 'fight':['fought','fighting'],
 'find':['found','finding'],
 'fly':['flew','flown','flying'],
 'forget':['forgot','forgotten','forgetting'],
 'freeze':['froze','frozen','freezing'],
 'get':['got','gotten','getting'],
 'give':['gave','given','giving'],
 'go':['goes','went','gone','going'],
 'grow':['grew','grown','growing'],
 'hang':['hung','hanging'],
 'have':['has','had','having'],
 'hear':['heard','hearing'],
 'hide':['hid','hidden','hiding'],
 'hold':['held','holding'],
 'keep':['kept','keeping'],
 'know':['knew','known','knowing'],
 'lay':['laid','laying'],
 'lead':['led','leading'],
 'leave':['left','leaving'],
 'lend':['lent','lending'],
 'let':['letting'],
 'lie':['lay','lain','lying'],
 'lose':['lost','losing'],
 'make':['made','making'],
 'mean':['meant','meaning'],
 'meet':['met','meeting'],
 'pay':['paid','paying'],
 'put':['putting'],
 'read':['read','reading'],
 'rid':['ridding'],
 'ride':['rode','ridden','riding'],
 'ring':['rang','rung','ringing'],
 'rise':['rose','risen','rising'],
 'run':['ran','running'],
 'say':['says','said','saying'],
 'see':['saw','seen','seeing'],
 'seek':['sought','seeking'],
 'sell':['sold','selling'],
 'send':['sent','sending'],
 'set':['setting'],
 'shake':['shook','shaken','shaking'],
 'shed':['shedding'],
 'shine':['shone','shining'],
 'shoot':['shot','shooting'],
 'show':['showed','shown','showing'],
 'shrink':['shrank','shrunk','shrinking'],
 'shut':['shutting'],
 'sing':['sang','sung','singing'],
 'sink':['sank','sunk','sinking'],
 'sit':['sat','sitting'],
 'sleep':['slept','sleeping'],
 'slide':['slid','sliding'],
 'speak':['spoke','spoken','speaking'],
 'spend':['spent','spending'],
 'spin':['spun','spinning'],
 'split':['splitting'],
 'spread':['spreading'],
 'spring':['sprang','sprung','springing'],
 'stand':['stood','standing'],
 'steal':['stole','stolen','stealing'],
 'stick':['stuck','sticking'],
 'sting':['stung','stinging'],
 'stink':['stank','stunk','stinking'],
 'strike':['struck','stricken','striking'],
 'swear':['swore','sworn','swearing'],
 'sweep':['swept','sweeping'],
 'swim':['swam','swum','swimming'],
 'swing':['swung','swinging'],
 'take':['took','taken','taking'],
 'teach':['taught','teaching'],
 'tear':['tore','torn','tearing'],
 'tell':['told','telling'],
 'think':['thought','thinking'],
 'throw':['threw','thrown','throwing'],
 'understand':['understood','understanding'],
 'wake':['woke','woken','waking'],
 'wear':['wore','worn','wearing'],
 'weave':['wove','woven','weaving'],
 'win':['won','winning'],
 'wind':['wound','winding'],
 'withdraw':['withdrew','withdrawn','withdrawing'],
 'write':['wrote','written','writing'],
}

POSSESSIVES = ["one's","his","her","its","their","my","your","our"]


def base_variants(w):
    forms = {w}
    forms.update([w+'s', w+'d', w+'ed', w+'ing'])
    if w.endswith('e'):
        forms.update([w[:-1]+'ing', w[:-1]+'ed'])
    if w.endswith('y') and len(w) > 1 and w[-2] not in 'aeiou':
        forms.update([w[:-1]+'ies', w[:-1]+'ied'])
    if len(w) >= 3 and w[-1] not in 'aeiouwxy' and w[-2] in 'aeiou' and w[-3] not in 'aeiou':
        forms.update([w+w[-1]+'ed', w+w[-1]+'ing'])
    if w in IRREGULAR:
        forms.update(IRREGULAR[w])
    if w.endswith(('s','x','z','ch','sh')):
        forms.add(w+'es')
    elif w.endswith('y') and len(w) > 1 and w[-2] not in 'aeiou':
        forms.add(w[:-1]+'ies')
    else:
        forms.add(w+'s')
    if w.endswith('e') and not w.endswith('le'):
        forms.add(w[:-1]+'ly')
    elif w.endswith('y') and len(w) > 1 and w[-2] not in 'aeiou':
        forms.add(w[:-1]+'ily')
    else:
        forms.add(w+'ly')
    forms.update('un'+f for f in list(forms))
    return forms


def token_alternatives(token):
    tl = token.lower()
    if tl == "one's":
        return POSSESSIVES
    if '/' in token:
        return [p.strip() for p in token.split('/') if p.strip()]
    return [token]


def find_blank(term, example):
    t = re.sub(r'\s*\([^)]*\)\s*$', '', term).strip()
    if not t:
        t = term.strip()
    tokens = t.split()
    n = len(tokens)
    expanded_token_choices = [token_alternatives(tok) for tok in tokens]

    candidates = []
    for combo in product(*expanded_token_choices):
        combo = list(combo)
        if n == 1:
            for v in base_variants(combo[0].lower()):
                candidates.append(v)
        else:
            first_variants = base_variants(combo[0].lower()) if '/' not in tokens[0] and tokens[0].lower() != "one's" else {combo[0]}
            last_variants = base_variants(combo[-1].lower()) if '/' not in tokens[-1] and tokens[-1].lower() != "one's" else {combo[-1]}
            middle = combo[1:-1]
            for fv in first_variants:
                for lv in last_variants:
                    candidates.append(' '.join([fv] + middle + [lv]))
            candidates.append(' '.join(combo))

    seen, ordered = set(), []
    for c in candidates:
        cl = c.lower()
        if cl not in seen:
            seen.add(cl)
            ordered.append(c)
    ordered.sort(key=len, reverse=True)

    for cand in ordered:
        pattern = re.compile(r'\b' + re.escape(cand) + r'\b', re.IGNORECASE)
        m = pattern.search(example)
        if m:
            return example[:m.start()] + "<span class='blank'>___________</span>" + example[m.end():]
    return None


def clean_group(g):
    g2 = re.sub(r'\s*\([^)]*\)\s*$', '', g).strip()
    return g2 if g2 else g.strip()


def first_token_grp(g):
    return g.split(';')[0].strip()


def slugify(s):
    s = s.lower()
    s = re.sub(r'[^a-z0-9]+', '-', s).strip('-')
    return s or 'misc'


def main():
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df.columns = ['Group', 'Vocab', 'Meaning', 'Example']
    df['Group'] = df['Group'].fillna('Additional Vocabulary')
    df = df.dropna(subset=['Vocab', 'Meaning', 'Example'])

    df['GroupClean'] = df['Group'].apply(clean_group)
    df['GroupKey'] = df['GroupClean'].apply(lambda g: first_token_grp(g).lower())

    groups, labels = {}, {}
    for key in df['GroupKey'].unique():
        groups[key] = slugify(key)
        cands = df.loc[df['GroupKey'] == key, 'GroupClean'].apply(first_token_grp)
        labels[key] = Counter(cands).most_common(1)[0][0]

    records = []
    for _, row in df.iterrows():
        term = str(row['Vocab']).strip()
        meaning = str(row['Meaning']).strip()
        example = str(row['Example']).strip()
        cat = groups[row['GroupKey']]
        qs = find_blank(term, example)
        if not qs:
            qs = "<span class='blank'>___________</span> — " + example
        records.append({"verb": term, "cat": cat, "meaning": meaning, "ex": example, "qs": qs})

    cats_meta = {slug: labels[key] for key, slug in groups.items()}

    verbs_json = json.dumps(records, ensure_ascii=False)
    cats_json = json.dumps(cats_meta, ensure_ascii=False)

    html = open(TEMPLATE_PATH, encoding='utf-8').read()
    html = html.replace('__CAT_LABELS__', cats_json)
    html = html.replace('__VERBS__', verbs_json)

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"Built {OUTPUT_PATH}: {len(records)} words across {len(cats_meta)} categories.")


if __name__ == "__main__":
    main()
