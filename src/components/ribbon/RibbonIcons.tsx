/**
 * Ribbon icons — original line icons drawn in the spirit of Office's ribbon
 * (24px grid, 1.5px stroke, currentColor). Accent strokes use --ribbon-accent.
 */

import type { ReactNode, SVGProps } from 'react';

export type RibbonIconName =
  | 'undo'
  | 'redo'
  | 'newComment'
  | 'deleteComment'
  | 'previous'
  | 'next'
  | 'showComments'
  | 'trackChanges'
  | 'accept'
  | 'reject'
  | 'reviewingPane'
  | 'markup'
  | 'table'
  | 'picture'
  | 'link'
  | 'tag'
  | 'pageBreak'
  | 'margins'
  | 'orientation'
  | 'size'
  | 'readMode'
  | 'printLayout'
  | 'ruler'
  | 'zoom'
  | 'onePage'
  | 'pageWidth'
  | 'find'
  | 'replace'
  | 'print'
  | 'save'
  | 'download'
  | 'pdf'
  | 'chevronDown'
  | 'check'
  | 'editing'
  | 'reviewing'
  | 'viewing'
  | 'insertRowAbove'
  | 'insertRowBelow'
  | 'insertColLeft'
  | 'insertColRight'
  | 'deleteTable'
  | 'merge'
  | 'split'
  | 'close'
  | 'wordCount'
  | 'header'
  | 'footer'
  | 'pageNumber'
  | 'open';

const ACCENT = 'var(--ribbon-accent, #185abd)';

const paths: Record<RibbonIconName, ReactNode> = {
  undo: <path d="M9 14 4 9l5-5M4 9h10.5a5.5 5.5 0 0 1 0 11H11" />,
  redo: <path d="m15 14 5-5-5-5M20 9H9.5a5.5 5.5 0 0 0 0 11H13" />,
  newComment: (
    <>
      <path d="M4 5.5A1.5 1.5 0 0 1 5.5 4h13A1.5 1.5 0 0 1 20 5.5v9a1.5 1.5 0 0 1-1.5 1.5H10l-4 4v-4h-.5A1.5 1.5 0 0 1 4 14.5z" />
      <path d="M12 7v6M9 10h6" stroke={ACCENT} />
    </>
  ),
  deleteComment: (
    <>
      <path d="M4 5.5A1.5 1.5 0 0 1 5.5 4h13A1.5 1.5 0 0 1 20 5.5v9a1.5 1.5 0 0 1-1.5 1.5H10l-4 4v-4h-.5A1.5 1.5 0 0 1 4 14.5z" />
      <path d="m9.5 7.5 5 5m0-5-5 5" stroke="#c42b1c" />
    </>
  ),
  previous: (
    <>
      <path d="M5 5.5A1.5 1.5 0 0 1 6.5 4h11A1.5 1.5 0 0 1 19 5.5v8a1.5 1.5 0 0 1-1.5 1.5H10l-3.5 3.5V15A1.5 1.5 0 0 1 5 13.5z" />
      <path d="m13.5 7-3 2.5 3 2.5" stroke={ACCENT} />
    </>
  ),
  next: (
    <>
      <path d="M5 5.5A1.5 1.5 0 0 1 6.5 4h11A1.5 1.5 0 0 1 19 5.5v8a1.5 1.5 0 0 1-1.5 1.5H10l-3.5 3.5V15A1.5 1.5 0 0 1 5 13.5z" />
      <path d="m10.5 7 3 2.5-3 2.5" stroke={ACCENT} />
    </>
  ),
  showComments: (
    <>
      <path d="M3 6.5A1.5 1.5 0 0 1 4.5 5h10A1.5 1.5 0 0 1 16 6.5v6a1.5 1.5 0 0 1-1.5 1.5H8l-3 3v-3h-.5A1.5 1.5 0 0 1 3 12.5z" />
      <path d="M18 9h1.5A1.5 1.5 0 0 1 21 10.5v6a1.5 1.5 0 0 1-1.5 1.5H19v3l-3-3h-4.5" />
    </>
  ),
  trackChanges: (
    <>
      <path d="M6 3.5h8l4 4v13H6z" />
      <path d="M14 3.5v4h4" />
      <path d="M9 12h6" stroke="#c42b1c" />
      <path d="M9 15.5h4" stroke={ACCENT} />
      <path d="m15.5 14.5 1 1-3 3H12.5v-1z" stroke={ACCENT} />
    </>
  ),
  accept: (
    <>
      <path d="M6 3.5h8l4 4v13H6z" />
      <path d="M14 3.5v4h4" />
      <path d="m9 14 2.2 2.2L15.5 11" stroke="#107c10" />
    </>
  ),
  reject: (
    <>
      <path d="M6 3.5h8l4 4v13H6z" />
      <path d="M14 3.5v4h4" />
      <path d="m9.5 11.5 5 5m0-5-5 5" stroke="#c42b1c" />
    </>
  ),
  reviewingPane: (
    <>
      <rect x="3.5" y="4.5" width="17" height="15" rx="1.5" />
      <path d="M9 4.5v15" />
      <path d="M5.5 8h2M5.5 11h2M5.5 14h2" stroke={ACCENT} />
    </>
  ),
  markup: (
    <>
      <path d="M5 6h14M5 10h9M5 14h14M5 18h7" />
      <path d="M2.5 9v6" stroke={ACCENT} />
    </>
  ),
  table: (
    <>
      <rect x="3.5" y="4.5" width="17" height="15" rx="1" />
      <path d="M3.5 9.5h17M3.5 14.5h17M9.5 4.5v15M15 4.5v15" />
      <path
        d="M3.5 9.5V5.5a1 1 0 0 1 1-1h15a1 1 0 0 1 1 1v4z"
        fill={ACCENT}
        stroke="none"
        opacity=".35"
      />
    </>
  ),
  picture: (
    <>
      <rect x="3.5" y="4.5" width="17" height="15" rx="1.5" />
      <circle cx="9" cy="9.5" r="1.75" stroke={ACCENT} />
      <path d="m3.5 17 5-4.5 3.5 3 3-2.5 5.5 4.5" />
    </>
  ),
  link: (
    <path d="M10 14a4 4 0 0 0 5.66 0l3-3a4 4 0 0 0-5.66-5.66l-1 1M14 10a4 4 0 0 0-5.66 0l-3 3A4 4 0 0 0 11 18.66l1-1" />
  ),
  tag: (
    <>
      <path d="M4 5.5v5.1a1.5 1.5 0 0 0 .44 1.06l7.9 7.9a1.5 1.5 0 0 0 2.12 0l5.1-5.1a1.5 1.5 0 0 0 0-2.12l-7.9-7.9A1.5 1.5 0 0 0 10.6 4H5.5A1.5 1.5 0 0 0 4 5.5z" />
      <circle cx="8.5" cy="8.5" r="1.25" stroke={ACCENT} />
    </>
  ),
  pageBreak: (
    <>
      <path d="M6 3v6.5h12V3M6 21v-6.5h12V21" />
      <path d="M3 12h2.5M8 12h2.5M13.5 12H16M18.5 12H21" stroke={ACCENT} />
    </>
  ),
  margins: (
    <>
      <path d="M6 3.5h12v17H6z" />
      <path d="M8.5 6.5h7v11h-7z" stroke={ACCENT} strokeDasharray="1.5 1.5" />
    </>
  ),
  orientation: (
    <>
      <path d="M4 7.5h10v13H4z" />
      <path d="M10 3.5h10v8" stroke={ACCENT} />
      <path d="m18 9.5 2 2 2-2" stroke={ACCENT} />
    </>
  ),
  size: (
    <>
      <path d="M7 3.5h10v17H7z" />
      <path d="M3.5 3.5v17M2.5 3.5h2M2.5 20.5h2" stroke={ACCENT} />
    </>
  ),
  readMode: (
    <>
      <path d="M12 6.5C10 5 7 4.5 3.5 5v13c3.5-.5 6.5 0 8.5 1.5 2-1.5 5-2 8.5-1.5V5C17 4.5 14 5 12 6.5z" />
      <path d="M12 6.5v13" stroke={ACCENT} />
    </>
  ),
  printLayout: (
    <>
      <path d="M6 3.5h12v17H6z" />
      <path d="M9 7.5h6M9 10.5h6M9 13.5h4" stroke={ACCENT} />
    </>
  ),
  ruler: (
    <>
      <rect x="2.5" y="8" width="19" height="8" rx="1" />
      <path d="M6 8v3M9.5 8v2M13 8v3M16.5 8v2M20 8v3" stroke={ACCENT} />
    </>
  ),
  zoom: (
    <>
      <circle cx="10.5" cy="10.5" r="6" />
      <path d="m15 15 5.5 5.5" />
      <path d="M8 10.5h5M10.5 8v5" stroke={ACCENT} />
    </>
  ),
  onePage: (
    <>
      <path d="M7 3.5h10v17H7z" />
      <path d="M12 7v10" stroke={ACCENT} strokeDasharray="1.5 1.5" />
    </>
  ),
  pageWidth: (
    <>
      <path d="M7 3.5h10v17H7z" />
      <path d="M2.5 12h19M4.5 10l-2 2 2 2M19.5 10l2 2-2 2" stroke={ACCENT} />
    </>
  ),
  find: (
    <>
      <circle cx="10.5" cy="10.5" r="6" />
      <path d="m15 15 5.5 5.5" stroke={ACCENT} />
    </>
  ),
  replace: (
    <>
      <path d="M4 7h11l-3-3M20 17H9l3 3" />
      <path
        d="M15 7h1.5A3.5 3.5 0 0 1 20 10.5V12M9 17H7.5A3.5 3.5 0 0 1 4 13.5V12"
        stroke={ACCENT}
      />
    </>
  ),
  print: (
    <>
      <path d="M7 8V3.5h10V8" />
      <rect x="3.5" y="8" width="17" height="8.5" rx="1.5" />
      <path d="M7 14h10v6.5H7z" stroke={ACCENT} />
    </>
  ),
  save: (
    <>
      <path d="M4.5 4.5h12l3 3v12h-15z" />
      <path d="M8 4.5v4.5h7V4.5M7.5 19.5v-6h9v6" stroke={ACCENT} />
    </>
  ),
  download: (
    <>
      <path d="M12 3.5v11M7.5 10l4.5 4.5 4.5-4.5" stroke={ACCENT} />
      <path d="M4 15.5v4h16v-4" />
    </>
  ),
  pdf: (
    <>
      <path d="M6 3.5h8l4 4v13H6z" />
      <path d="M14 3.5v4h4" />
      <path
        d="M8.5 16.5v-4h1.25a1.25 1.25 0 0 1 0 2.5H8.5M12.5 16.5v-4h.75a2 2 0 0 1 0 4zM16.5 12.5h-1.5v4M15 14.5h1.25"
        stroke="#c42b1c"
        strokeWidth="1.1"
      />
    </>
  ),
  chevronDown: <path d="m7 10 5 5 5-5" />,
  check: <path d="m5 12.5 4.5 4.5L19 7.5" />,
  editing: (
    <>
      <path d="m14.5 5.5 4 4L8 20H4v-4z" />
      <path d="m12.5 7.5 4 4" stroke={ACCENT} />
    </>
  ),
  reviewing: (
    <>
      <path d="M4 5.5A1.5 1.5 0 0 1 5.5 4h13A1.5 1.5 0 0 1 20 5.5v9a1.5 1.5 0 0 1-1.5 1.5H10l-4 4v-4h-.5A1.5 1.5 0 0 1 4 14.5z" />
      <path d="m14 7 2 2-5 5H9v-2z" stroke={ACCENT} />
    </>
  ),
  viewing: (
    <>
      <path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z" />
      <circle cx="12" cy="12" r="3" stroke={ACCENT} />
    </>
  ),
  insertRowAbove: (
    <>
      <rect x="3.5" y="11" width="17" height="9" rx="1" />
      <path d="M3.5 15.5h17M9.5 11v9M15 11v9" />
      <path d="M12 3v6M9 6h6" stroke={ACCENT} />
    </>
  ),
  insertRowBelow: (
    <>
      <rect x="3.5" y="4" width="17" height="9" rx="1" />
      <path d="M3.5 8.5h17M9.5 4v9M15 4v9" />
      <path d="M12 15v6M9 18h6" stroke={ACCENT} />
    </>
  ),
  insertColLeft: (
    <>
      <rect x="11" y="3.5" width="9" height="17" rx="1" />
      <path d="M15.5 3.5v17M11 9.5h9M11 15h9" />
      <path d="M3 12h6M6 9v6" stroke={ACCENT} />
    </>
  ),
  insertColRight: (
    <>
      <rect x="4" y="3.5" width="9" height="17" rx="1" />
      <path d="M8.5 3.5v17M4 9.5h9M4 15h9" />
      <path d="M15 12h6M18 9v6" stroke={ACCENT} />
    </>
  ),
  deleteTable: (
    <>
      <rect x="3.5" y="4.5" width="13" height="11" rx="1" />
      <path d="M3.5 10h13M10 4.5v11" />
      <path d="m15 15 5 5m0-5-5 5" stroke="#c42b1c" />
    </>
  ),
  merge: (
    <>
      <rect x="3.5" y="5.5" width="17" height="13" rx="1" />
      <path d="M7 12h10M9 9.5 6.5 12 9 14.5M15 9.5l2.5 2.5-2.5 2.5" stroke={ACCENT} />
    </>
  ),
  split: (
    <>
      <rect x="3.5" y="5.5" width="17" height="13" rx="1" />
      <path d="M12 5.5v13" />
      <path d="M5.5 12h4.5M14 12h4.5" stroke={ACCENT} />
    </>
  ),
  close: <path d="m6 6 12 12M18 6 6 18" />,
  header: (
    <>
      <path d="M6 3.5h12v17H6z" />
      <path d="M6 8h12" />
      <path d="M8.5 5.75h7" stroke={ACCENT} />
      <path d="M8.5 11h7M8.5 14h7M8.5 17h4" strokeOpacity=".45" />
    </>
  ),
  footer: (
    <>
      <path d="M6 3.5h12v17H6z" />
      <path d="M6 16h12" />
      <path d="M8.5 18.25h7" stroke={ACCENT} />
      <path d="M8.5 7h7M8.5 10h7M8.5 13h4" strokeOpacity=".45" />
    </>
  ),
  open: (
    <>
      <path d="M3.5 6.5A1.5 1.5 0 0 1 5 5h4.5l2 2H19a1.5 1.5 0 0 1 1.5 1.5V10" />
      <path d="M3.5 6.5v11.5h14l3-8H6.5l-3 8" stroke={ACCENT} />
    </>
  ),
  pageNumber: (
    <>
      <path d="M6 3.5h12v17H6z" />
      <path d="M8.5 7h7M8.5 10h7M8.5 13h4" strokeOpacity=".45" />
      <path d="M11 16.5l1.25-.75V19.5" stroke={ACCENT} />
    </>
  ),
  wordCount: (
    <>
      <path d="M5 6h10M5 10h14M5 14h8" />
      <path d="M16 15.5h4M16 18.5h4M17.5 14v6M18.5 14v6" stroke={ACCENT} strokeWidth="1.1" />
    </>
  ),
};

export interface RibbonIconProps extends Omit<SVGProps<SVGSVGElement>, 'name'> {
  name: RibbonIconName;
  size?: number;
}

export function RibbonIcon({ name, size = 20, ...rest }: RibbonIconProps) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.5}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
      {...rest}
    >
      {paths[name]}
    </svg>
  );
}
