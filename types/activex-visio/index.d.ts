/// <reference types="activex-office" />
/// <reference types="activex-vbide" />
/// <reference types="activex-stdole" />
/// <reference types="activex-adodb" />
/// <reference types="activex-dao" />

declare namespace Visio {

    class Application {
        private "Visio.Application_typekey": Application;
        private constructor();
        readonly _Default: string;
        readonly Application: Application;
        readonly ActiveDocument: Document;
        readonly ActivePage: Page;
        ActivePrinter: string;
        readonly ActiveWindow: Window;
        readonly Name: string;

    }

    class Document {
        private "Visio.Document_typekey": Document;
        private constructor();
    }

    interface Documents {
        /** @param TextQualifier [TextQualifier=1] */
    }

    class Page {
        private "Visio.Page_typekey": Page;
        private constructor();
       }

    interface ActiveXObjectNameMap {
        "Visio.Application": Visio.Application;
        "Visio.Document": Visio.Document;
        "Visio.Page": Visio.Page;
    }
   
}
