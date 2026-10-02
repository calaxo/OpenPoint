// ═══ [PPTX · UNITÉS] ═══ EMU -> points.
// EMU -> points (1 pt = 12700 EMU). On rend en points pour des tailles de police normales
// (les foreignObject cassent avec des font-size en centaines de milliers de px).
// S = facteur EMU -> pt. u(v) = valeur en pt arrondie au 1/100 (toFixed(2) puis + pour revenir à un
// nombre : pas de « 12.50 » dans la sortie, mais « 12.5 »). Cet arrondi est appliqué PARTOUT où une
// coordonnée sort dans le SVG : le changer modifierait tous les rendus (empreintes). Note : u()
// renvoie un nombre ; plusieurs appels font ensuite .toFixed(1) dessus pour un format plus court.
const S = 1 / 12700;
const u = (v) => +(Number(v) * S).toFixed(2);
export { S, u };
