# Third-party components

Margin's original code is licensed under the MIT License in `LICENSE`. That license does not replace any third-party component's license.

| Component | Terms | Included material |
| --- | --- | --- |
| `lib/vendor/latexdiff-so` | GPL-3.0-or-later, copyright Frederik Tilmann and contributors | Complete, unmodified standalone Perl program; its notices remain in the source and GPL text in `lib/vendor/COPYING`. Invoked as a separate command-line tool to compare two LaTeX files. |
| `diff` | BSD-3-Clause, copyright Kevin Decker | Installed through npm; retain its package license when redistributing the dependency. |
| Paper Mono | SIL Open Font License 1.1 | Variable font and complete license in `public/fonts/PAPER-MONO-OFL.txt`. Upstream: [paper-design/paper-mono](https://github.com/paper-design/paper-mono), commit `e6eaeceaef02e77e3db997711e07a16378de2bd7`. |

The GPL component is supplied under its own terms, including its complete source, rather than relicensed as MIT. MIT is GPL-compatible. The programs communicate through ordinary command-line arguments and LaTeX files; the [GNU license FAQ on aggregation](https://www.gnu.org/licenses/gpl-faq.html#MereAggregation) explains the distinction between separate programs distributed together and a combined program. Distributors must preserve the rights and notices of each included component; an MIT notice does not remove GPL obligations for `latexdiff-so`.

Pandoc, TeX Live/MacTeX, Poppler, Python, lxml and pypdf are separately installed dependencies. Their respective licenses apply to those installations; their binaries are not bundled here.

Timeless is an optional local font and is **not included** in this repository. Its distributor labels it free, but the obtained archive did not contain an explicit redistribution license. Users may obtain it independently and should review the applicable terms. No OFL or MIT claim is made for Timeless.
