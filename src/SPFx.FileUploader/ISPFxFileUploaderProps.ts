
import { SPHttpClient } from '@microsoft/sp-http';
import { SPWeb } from '@microsoft/sp-page-context';


interface ISPFxFileUploaderProps {
    /** Returns the files added if want to have custom logic while uplaoding */
    onFileAdded?: (files: ISPFxFile[] ) => void;

    /** If want to use the component upload functionality */
    enableUpload: boolean;

    spWeb?: SPWeb;

    docLibInternalName?: string;

    folder?: string;

    spHttp: SPHttpClient;
}

interface ISPFxFileUploaderState {
    selectedFiles?: ISPFxFile[];
    spoFiles?: string[];
}

interface ISPFxFile {
    file: File;
    fileState: IFileState;
}

enum IFileState {
    READY,
    UPLOADED,
    ERROR
}

export { ISPFxFileUploaderProps };
export { ISPFxFileUploaderState };
export { ISPFxFile };
export { IFileState };