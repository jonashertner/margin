export const before=String.raw`\documentclass[11pt,a4paper]{article}
\usepackage[margin=28mm]{geometry}
\usepackage{fontspec}
\setmainfont{texgyrepagella}[Extension=.otf,UprightFont=*-regular,BoldFont=*-bold,ItalicFont=*-italic,BoldItalicFont=*-bolditalic]
\usepackage{parskip}
\usepackage[hidelinks]{hyperref}
\hypersetup{pdfauthor={},pdftitle={},pdfsubject={},pdfkeywords={}}
\title{Settlement agreement}
\author{}
\date{}
\begin{document}
\maketitle

\noindent\textit{Fictional drafting example. Not for execution.}

\section{Parties and purpose}

This agreement is made between Aster AG (the Company) and Linden GmbH (the Contractor), together the Parties, to resolve their dispute concerning the services agreement dated 12 January 2026.

\section{Settlement payment}

The Company shall pay the Contractor CHF 125,000 in full and final settlement of the claims described in clause 4.

Payment shall be made within 30 days of execution of this agreement to the account designated in writing by the Contractor.

Each Party shall bear its own legal costs in connection with the dispute and this agreement.

\section{Confidentiality}

The Parties shall keep the existence and terms of this agreement confidential, except where disclosure is required by law.

\section{Release of claims}

Upon execution of this agreement, each Party irrevocably releases the other from all claims arising out of the services agreement.

The release does not affect the obligations created by this agreement.

\section{Governing law and jurisdiction}

This agreement is governed by Swiss law. The courts of Zurich shall have exclusive jurisdiction over any dispute arising out of this agreement.

\section{Execution}

This agreement may be executed in counterparts. Each counterpart shall constitute an original.

\bigskip
\noindent For Aster AG \hfill For Linden GmbH

\vspace{18mm}
\noindent Signature: \hrulefill\hspace{12mm} Signature: \hrulefill

\end{document}
`;
export const after=before
.replace('CHF 125,000','CHF 150,000')
.replace('within 30 days of execution','within 15 business days of execution')
.replace('except where disclosure is required by law.','except where disclosure is required by law or to professional advisers who are bound by a duty of confidentiality.')
.replace('Upon execution of this agreement, each Party irrevocably releases','Upon receipt of the settlement payment in cleared funds, each Party releases')
.replace('The release does not affect the obligations created by this agreement.','The release does not affect the obligations created by this agreement or any claim arising from fraud or wilful misconduct.')
.replace('This agreement may be executed in counterparts. Each counterpart shall constitute an original.','This agreement may be executed in counterparts and signed electronically. Each counterpart shall constitute an original.');
export const demo={title:'Settlement agreement',matter:'Aster AG / Linden GmbH',before,after,beforeName:'Settlement_v1.tex',afterName:'Settlement_v2.tex',decisions:{},comments:{},audit:[],demo:true};
