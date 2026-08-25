/**
 * Vendored from office-word-diff (Apache-2.0) — see LICENSE and NOTICE in
 * this directory. Implements the word-mode diff method from Google's
 * diff-match-patch library, extending the prototype safely (without editing
 * node_modules).
 *
 * @module lib/word-diff/diff-wordmode
 */

import DiffMatchPatch from 'diff-match-patch';

DiffMatchPatch.prototype.diff_linesToWords_ = function (a, b) {
    function c(a) {
        var re = /(\w+|[^\w\s]+|\s+)/g; // split into words, punctuation, and whitespace
        var b = '';
        var match;
        while ((match = re.exec(a)) !== null) {
            var token = match[0];
            if (Object.prototype.hasOwnProperty.call(e, token)) {
                b += String.fromCharCode(e[token]);
            } else {
                d.push(token);
                e[token] = d.length - 1;
                b += String.fromCharCode(e[token]);
            }
        }
        return b;
    }

    var d = [], e = {};
    d[0] = '';
    var g = c(a);
    var h = c(b);
    return { chars1: g, chars2: h, lineArray: d };
};

DiffMatchPatch.prototype.diff_wordMode = function (text1, text2) {
    const dmp = this;
    const a = dmp.diff_linesToWords_(text1, text2);
    const wordText1 = a.chars1;
    const wordText2 = a.chars2;
    const wordArray = a.lineArray;

    const diffs = dmp.diff_main(wordText1, wordText2, false);
    dmp.diff_charsToLines_(diffs, wordArray);
    // dmp.diff_cleanupSemantic(diffs); // Disabled to preserve token boundaries
    return diffs;
};

export default DiffMatchPatch;
