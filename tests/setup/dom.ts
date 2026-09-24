// The core parses field SchemaXml with the browser's DOMParser / XMLSerializer. Tests run in the Node
// environment (for fetch/Response), so only these two globals are borrowed from jsdom.
import { JSDOM } from 'jsdom';

const { window } = new JSDOM('');
Object.assign(globalThis, { DOMParser: window.DOMParser, XMLSerializer: window.XMLSerializer });
