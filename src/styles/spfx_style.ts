export const HIDE_SWITCHER_CSS = `

button[data-automationid="librariesDropdownButton"] {
  display: none !important;
}

button[title='Switch to other libraries on this site'] {
    display: none !important;
}

button[title='Alternar para outras bibliotecas neste site'] {
    display: none !important;
}

.librariesDropdown_7ab24cf7{
    display: none !important;
}

div[data-automationid="commandbar"] {
  display: none !important;
}

div[data-automationid="appCommandBar"] {
  height: 0 !important;
  min-height: 0 !important;
  overflow: hidden !important;
}

button[data-automationid="filterPill"] i[data-icon-name="Cancel"] {
  display: none !important;
}

button[data-automationid="clearFiltersPill"] {
  display: none !important;
}

div[data-automationid^="addColumnLollipop"] {
  display: none !important;
  pointer-events: none !important;
}

[class^="resizeDivider"],
[class*=" resizeDivider"] {
  display: none !important;
  pointer-events: none !important;
}

[class^="headerCell"],
[class*=" headerCell"] {
  cursor: default !important;
  resize: none !important;
}

.splitter-layout .layout-pane.layout-pane-primary {
  overflow: hidden !important;
  flex: 1 1 auto !important;
  display: flex !important;
  flex-direction: column !important;
}

.splitter-layout .layout-pane.layout-pane-primary .iframeContainer_b3e2de1e {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.splitter-layout .layout-pane.layout-pane-primary .iframeContainer_b3e2de1e > div {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.ms-FlowPanel-contents iframe {
  width: 100%;
  height: 80% !important;
  border: 0;
}

`;
