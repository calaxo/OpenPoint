// Tests unitaires des fonctions PURES de tools/parse-timing.mjs.
// Cœur métier : la lecture de l'arbre <p:timing> d'un PPTX (format "preserveOrder"
// de fast-xml-parser) et sa transformation en timeline JSON. On ne teste QUE des
// fonctions dont la sortie ne dépend que des entrées (aucun fs/réseau/DOM).
// ESM (test/package.json => "type": "module"), avec node:test + node:assert/strict.
//
// --- Pourquoi un import() dynamique plutôt que statique ? ---
// Le module cible est un ESM qui, à son chargement, exécute un bloc "main" lisant
// process.argv et faisant des I/O. On prépare donc argv pour l'orienter vers un
// chemin INOFFENSIF — parseSlide(1), qui lit slide1.xml (présent), logge et ne throw
// pas, sans rien réécrire — AVANT de charger le module. Un import statique aurait
// tourné avant qu'on ait pu ajuster argv ; le import() dynamique nous rend la main.
import test from 'node:test';
import assert from 'node:assert/strict';

process.argv = [process.argv[0], 'parse-timing.mjs', '1'];
const _log = console.log;
console.log = () => {}; // on masque le console.log du bloc "main" pendant le chargement
const load = import('../tools/parse-timing.mjs').finally(() => {
  console.log = _log;
});

// ---------------------------------------------------------------------------
// Helpers d'arbre : tagOf / attrs / kids
// Un noeud preserveOrder = { <tagName>: [enfants...], ':@': { attributs } }.
// Un noeud texte = { '#text': 'valeur' }.
// ---------------------------------------------------------------------------
test('tagOf renvoie le nom de balise (clé hors ":@")', async () => {
  const { tagOf } = await load;
  assert.equal(tagOf({ 'p:cTn': [], ':@': { '@id': '1' } }), 'p:cTn');
  assert.equal(tagOf({ '#text': 'bonjour' }), '#text');
});

test('attrs renvoie les attributs, ou {} si absent', async () => {
  const { attrs } = await load;
  assert.deepEqual(attrs({ 'p:sp': [], ':@': { '@spid': '9' } }), { '@spid': '9' });
  assert.deepEqual(attrs({ 'p:sp': [] }), {});
});

test('kids renvoie le tableau d\'enfants, ou [] si absent', async () => {
  const { kids } = await load;
  const enfant = { '#text': 'x' };
  assert.deepEqual(kids({ 'p:cTn': [enfant] }), [enfant]);
  assert.deepEqual(kids({ ':@': { '@a': '1' } }), []); // pas de balise -> []
});

// ---------------------------------------------------------------------------
// childrenByTag / firstChild : sélection d'enfants par nom de balise
// ---------------------------------------------------------------------------
test('childrenByTag ne garde que les enfants du tag demandé (ordre préservé)', async () => {
  const { childrenByTag } = await load;
  const node = {
    'p:childTnLst': [
      { 'p:par': [], ':@': { '@n': '1' } },
      { 'p:set': [] },
      { 'p:par': [], ':@': { '@n': '2' } },
    ],
  };
  const pars = childrenByTag(node, 'p:par');
  assert.equal(pars.length, 2);
  assert.equal(pars[0][':@']['@n'], '1');
  assert.equal(pars[1][':@']['@n'], '2');
  assert.deepEqual(childrenByTag(node, 'p:absent'), []);
});

test('firstChild renvoie le premier enfant du tag, sinon undefined', async () => {
  const { firstChild } = await load;
  const node = { 'p:cTn': [{ 'p:set': [], ':@': { '@k': 'a' } }, { 'p:set': [], ':@': { '@k': 'b' } }] };
  assert.equal(firstChild(node, 'p:set')[':@']['@k'], 'a');
  assert.equal(firstChild(node, 'p:cmd'), undefined);
});

// ---------------------------------------------------------------------------
// findFirst : recherche récursive en profondeur (renvoie le noeud lui-même
// s'il correspond, sinon le premier descendant, sinon null)
// ---------------------------------------------------------------------------
test('findFirst : match sur le noeud racine', async () => {
  const { findFirst } = await load;
  const n = { 'p:cond': [], ':@': { '@delay': '0' } };
  assert.equal(findFirst(n, 'p:cond'), n);
});

test('findFirst : descend chercher un noeud imbriqué', async () => {
  const { findFirst } = await load;
  const tree = {
    'p:cBhvr': [
      { 'p:cTn': [{ 'p:stCondLst': [{ 'p:cond': [], ':@': { '@delay': '500' } }] }] },
    ],
  };
  const cond = findFirst(tree, 'p:cond');
  assert.ok(cond);
  assert.equal(cond[':@']['@delay'], '500');
});

test('findFirst : null quand rien ne correspond', async () => {
  const { findFirst } = await load;
  assert.equal(findFirst({ 'p:sp': [{ '#text': 'x' }] }, 'p:introuvable'), null);
});

// ---------------------------------------------------------------------------
// shapeMap : parse le XML brut de la diapo -> { spid: nom de forme }
// (fonction pure sur une string ; regex globale sur <p:cNvPr id=".." name="..">)
// ---------------------------------------------------------------------------
test('shapeMap extrait id -> name pour chaque p:cNvPr', async () => {
  const { shapeMap } = await load;
  const xml =
    '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Titre 1"/></p:nvSpPr></p:sp>' +
    '<p:pic><p:nvPicPr><p:cNvPr id="7" name="Image 2"/></p:nvPicPr></p:pic>';
  assert.deepEqual(shapeMap(xml), { '4': 'Titre 1', '7': 'Image 2' });
});

test('shapeMap sur un XML sans forme renvoie {}', async () => {
  const { shapeMap } = await load;
  assert.deepEqual(shapeMap('<p:timing></p:timing>'), {});
});

// ---------------------------------------------------------------------------
// parseBehavior : extrait un behavior (p:set, p:anim, p:animMotion, ...)
// -> { kind, attributs, dur/fill/repeat, delay, spid, keyframes, ... }
// ---------------------------------------------------------------------------
test('parseBehavior (p:set) : attributs, cBhvr, cible, plage de paragraphes, valeur "to"', async () => {
  const { parseBehavior } = await load;
  const setNode = {
    'p:set': [
      {
        'p:cBhvr': [
          {
            'p:cTn': [
              { 'p:stCondLst': [{ 'p:cond': [], ':@': { '@delay': '500' } }] },
            ],
            ':@': { '@dur': '1000', '@fill': 'hold', '@repeatCount': '2', '@autoRev': '1' },
          },
          {
            'p:tgtEl': [
              {
                'p:spTgt': [{ 'p:txEl': [{ 'p:pRg': [], ':@': { '@st': '0', '@end': '2' } }] }],
                ':@': { '@spid': '42' },
              },
            ],
          },
          { 'p:attrNameLst': [{ 'p:attrName': [{ '#text': 'style.visibility' }] }] },
        ],
      },
      { 'p:to': [{ 'p:strVal': [], ':@': { '@val': 'visible' } }] },
    ],
    ':@': { '@filter': 'wipe(right)', '@transition': 'in' },
  };
  const b = parseBehavior(setNode);
  assert.equal(b.kind, 'set');
  assert.equal(b.filter, 'wipe(right)');
  assert.equal(b.transition, 'in');
  assert.equal(b.dur, '1000');
  assert.equal(b.fill, 'hold');
  assert.equal(b.repeat, '2');
  assert.equal(b.autoRev, '1');
  assert.equal(b.delay, '500');
  assert.equal(b.spid, '42');
  assert.deepEqual(b.paragraphRange, ['0', '2']);
  assert.deepEqual(b.attrs, ['style.visibility']);
  assert.equal(b.to, 'visible');
});

test('parseBehavior (p:anim) : keyframes tavLst (tm/val/fmla)', async () => {
  const { parseBehavior } = await load;
  const animNode = {
    'p:anim': [
      {
        'p:cBhvr': [
          { 'p:cTn': [], ':@': {} },
          { 'p:tgtEl': [{ 'p:spTgt': [], ':@': { '@spid': '9' } }] },
        ],
      },
      {
        'p:tavLst': [
          { 'p:tav': [{ 'p:val': [{ 'p:fltVal': [], ':@': { '@val': '0' } }] }], ':@': { '@tm': '0' } },
          {
            'p:tav': [{ 'p:val': [{ 'p:fltVal': [], ':@': { '@val': '1' } }] }],
            ':@': { '@tm': '100000', '@fmla': '#ppt_x+0.5' },
          },
        ],
      },
    ],
    ':@': { '@calcmode': 'lin', '@valueType': 'num' },
  };
  const b = parseBehavior(animNode);
  assert.equal(b.kind, 'anim');
  assert.equal(b.calcmode, 'lin');
  assert.equal(b.valueType, 'num');
  assert.equal(b.spid, '9');
  assert.deepEqual(b.keyframes, [
    { tm: '0', val: '0' },
    { tm: '100000', fmla: '#ppt_x+0.5', val: '1' },
  ]);
});

test('parseBehavior (p:animMotion) : chemin + valeur "by" en {x,y}', async () => {
  const { parseBehavior } = await load;
  const motionNode = {
    'p:animMotion': [
      { 'p:cBhvr': [{ 'p:cTn': [], ':@': {} }, { 'p:tgtEl': [{ 'p:spTgt': [], ':@': { '@spid': '3' } }] }] },
      { 'p:by': [{ 'p:pt': [], ':@': { '@x': '10', '@y': '20' } }] },
    ],
    ':@': { '@path': 'M 0 0 L 1 1' },
  };
  const b = parseBehavior(motionNode);
  assert.equal(b.kind, 'animMotion');
  assert.equal(b.path, 'M 0 0 L 1 1');
  assert.deepEqual(b.by, { x: '10', y: '20' });
  assert.equal(b.spid, '3');
});

// ---------------------------------------------------------------------------
// parseEffect : un <p:par> porteur de presetID -> descripteur d'effet.
// Le délai effectif = accDelay (cumul des groupes parents) + délai propre.
// ---------------------------------------------------------------------------
test('parseEffect : preset, spid, forme résolue, délai cumulé, behaviors', async () => {
  const { parseEffect } = await load;
  const parNode = {
    'p:par': [
      {
        'p:cTn': [
          { 'p:stCondLst': [{ 'p:cond': [], ':@': { '@delay': '200' } }] },
          {
            'p:childTnLst': [
              {
                'p:set': [
                  {
                    'p:cBhvr': [
                      { 'p:cTn': [], ':@': {} },
                      { 'p:tgtEl': [{ 'p:spTgt': [], ':@': { '@spid': '7' } }] },
                    ],
                  },
                ],
              },
            ],
          },
        ],
        ':@': { '@nodeType': 'clickEffect', '@presetClass': 'entr', '@presetID': '1', '@presetSubtype': '0' },
      },
    ],
  };
  const eff = parseEffect(parNode, { '7': 'Rectangle 3' }, 1000);
  assert.equal(eff.nodeType, 'clickEffect');
  assert.equal(eff.preset, 'entr:1:0');
  assert.equal(eff.spid, '7');
  assert.equal(eff.shape, 'Rectangle 3');
  assert.equal(eff.delay, 1200); // 1000 (parents) + 200 (propre)
  assert.equal(eff.behaviors.length, 1);
  assert.equal(eff.behaviors[0].kind, 'set');
  assert.equal(eff.behaviors[0].spid, '7');
});

test('parseEffect : presetSubtype absent -> défaut 0, forme inconnue -> null', async () => {
  const { parseEffect } = await load;
  const parNode = {
    'p:par': [
      {
        'p:cTn': [{ 'p:childTnLst': [] }],
        ':@': { '@nodeType': 'clickEffect', '@presetClass': 'emph', '@presetID': '5' },
      },
    ],
  };
  const eff = parseEffect(parNode, {}, 0);
  assert.equal(eff.preset, 'emph:5:0'); // subtype par défaut
  assert.equal(eff.spid, undefined);
  assert.equal(eff.shape, null);
  assert.equal(eff.delay, 0); // pas de stCondLst -> ownDelay 0
  assert.deepEqual(eff.behaviors, []);
});

// ---------------------------------------------------------------------------
// collectEffects : traverse un click-step et accumule les délais des groupes
// de synchro parents. Un délai "indefinite" (le clic lui-même) n'est pas cumulé.
// ---------------------------------------------------------------------------
test('collectEffects : délai final = somme des groupes + délai propre de l\'effet', async () => {
  const { collectEffects } = await load;
  const tree = {
    'p:par': [ // niveau clic : delay=indefinite (ne s'ajoute pas)
      {
        'p:cTn': [
          { 'p:stCondLst': [{ 'p:cond': [], ':@': { '@delay': 'indefinite' } }] },
          {
            'p:childTnLst': [
              {
                'p:par': [ // groupe de synchro : delay=100 (s'ajoute)
                  {
                    'p:cTn': [
                      { 'p:stCondLst': [{ 'p:cond': [], ':@': { '@delay': '100' } }] },
                      {
                        'p:childTnLst': [
                          {
                            'p:par': [ // effet : delay propre=50, presetID présent
                              {
                                'p:cTn': [
                                  { 'p:stCondLst': [{ 'p:cond': [], ':@': { '@delay': '50' } }] },
                                  {
                                    'p:childTnLst': [
                                      {
                                        'p:set': [
                                          {
                                            'p:cBhvr': [
                                              { 'p:cTn': [], ':@': {} },
                                              { 'p:tgtEl': [{ 'p:spTgt': [], ':@': { '@spid': '5' } }] },
                                            ],
                                          },
                                        ],
                                      },
                                    ],
                                  },
                                ],
                                ':@': { '@nodeType': 'afterEffect', '@presetClass': 'emph', '@presetID': '3' },
                              },
                            ],
                          },
                        ],
                      },
                    ],
                    ':@': {},
                  },
                ],
              },
            ],
          },
        ],
        ':@': {},
      },
    ],
  };
  const acc = [];
  collectEffects(tree, { '5': 'Star 1' }, acc);
  assert.equal(acc.length, 1);
  const eff = acc[0];
  assert.equal(eff.preset, 'emph:3:0');
  assert.equal(eff.spid, '5');
  assert.equal(eff.shape, 'Star 1');
  assert.equal(eff.delay, 150); // 100 (groupe) + 50 (propre) ; le clic "indefinite" ignoré
  assert.equal(eff.nodeType, 'afterEffect');
});

test('collectEffects : plusieurs effets frères collectés dans l\'ordre document', async () => {
  const { collectEffects } = await load;
  const mkEffet = (id, spid) => ({
    'p:par': [
      {
        'p:cTn': [{ 'p:childTnLst': [] }],
        ':@': { '@nodeType': 'clickEffect', '@presetClass': 'entr', '@presetID': id, '@presetSubtype': '0' },
      },
    ],
  });
  const groupe = {
    'p:par': [
      {
        'p:cTn': [{ 'p:childTnLst': [mkEffet('1'), mkEffet('2')] }],
        ':@': {},
      },
    ],
  };
  const acc = [];
  collectEffects(groupe, {}, acc);
  assert.deepEqual(acc.map((e) => e.preset), ['entr:1:0', 'entr:2:0']);
});
