import * as React from 'react';
// import styles from './SPFxFileUploader.module.scss';
const styles = require('./SPFxFileUploader.css');
import { ISPFxFileUploaderProps, ISPFxFileUploaderState, ISPFxFile, IFileState } from './ISPFxFileUploaderProps';
import { Button, ButtonType, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
import { Log } from '@microsoft/sp-core-library';

const LOG_SOURCE: string = 'TOPSOE.FileUploader';

class SPFxFileUploader extends React.Component<ISPFxFileUploaderProps, ISPFxFileUploaderState>{

    constructor(props: ISPFxFileUploaderProps, state: ISPFxFileUploaderState) {
        super(props);
        this.state = {
            selectedFiles: [],
            spoFiles: []
        };
    }

    /**
     * returns the folder path for the files to be accessed
     */
    private get folderPath(): string {
        return `${this.props.docLibInternalName.replace(/[\/]+$/, '')}/${this.props.folder.replace(/[\/]+$/, '')}`;
    }

    /**
     * Returns the color of the round agains each file status
     * @param fileState 
     */
    private fileStateColor(fileState: IFileState): string {
        return fileState == IFileState.READY ? `#ffb100` : fileState == IFileState.UPLOADED ? `#008000` : `#ff0000`;
    }

    /**
     * Returns the state text agains each file
     * @param fileState 
     */
    private fileStateText(fileState: IFileState): string {
        return fileState == IFileState.READY ? `Ready for uplaod` : fileState == IFileState.UPLOADED ? `Uplaoded Successfully` : `Some error occured while uplaoding`;
    }

    public componentDidMount(): void {

        this.registerDragEvents();
        this.getFilesOnSpo();
    }

    /**
     * Registers the drag events of the file
     */
    private registerDragEvents() {
        let element = document.getElementById("spfxFileUploader");

        window.addEventListener("drop", (event) => {
            event.preventDefault();
        });

        /* Events fired on the drop target */
        element.addEventListener("dragover", (event) => {
            event.preventDefault();
        });

        element.addEventListener("drop", (event) => {

            console.log(event.dataTransfer.files);
            this.onFileUploaded(event.dataTransfer.files);
        });
    }

    public render(): React.ReactElement<ISPFxFileUploaderProps> {
        return (
            <div className={`SPFxFileUploader`} id={"spfxFileUploader"}>
                <div className={`container ms-Grid`}>
                    <div className={`uploadbtn ms-Grid-row ms-Grid-col ms-sm12`}>
                        <input type="file" id="fileUploader" multiple={true} name="file" onChange={(event) => this.onFileUploaded(event.target.files)}></input>
                        <PrimaryButton className={`addIcon`} iconProps={{ iconName: "Add" }} onClick={() => { document.getElementById('fileUploader').click(); }}>Add</PrimaryButton>
                        {
                            this.props.enableUpload ?
                                <PrimaryButton iconProps={{ iconName: "CloudUpload" }} onClick={() => this.uploadFilesToSP()} disabled={this.state.selectedFiles.length == 0}>Upload</PrimaryButton>
                                :
                                <div></div>
                        }
                    </div>
                    <div className={`showFiles ms-Grid-row ms-Grid-col ms-sm12`}>
                        {this.renderFiles()}
                    </div>
                </div>
            </div>
        );
    }


    /**
     * Call back of the files that needs to be uplaoded
     * @param event 
     */
    private onFileUploaded(fileList: FileList): void {

        let files: ISPFxFile[] = this.state.selectedFiles;
        let fileCount: number = 0;

        // Iterate over each file uploaded and asdd it to the state
        while (fileList.length > fileCount) {

            files.push({ file: fileList[fileCount], fileState: IFileState.READY });
            fileCount += 1;
        }

        this.setState({ selectedFiles: files }, () => {
            this.getFilesOnSpo();
            if (this.props.onFileAdded) this.props.onFileAdded(this.state.selectedFiles);
        });
    }


    /**
     * Renders HTML part fpr the files uploaded
     */
    private renderFiles(): JSX.Element {
        return (
            <div>
                <div className={`addFilesSection`}>
                    {
                        this.state.selectedFiles.length == 0 ?
                            <Label className={`noFileText`}>Add your files here</Label>
                            :
                            this.state.selectedFiles.map(item => {
                                return (<Label key={item.file.name} title={this.fileStateText(item.fileState)} className={`fileName`}>{item.file.name}
                                    <Icon className={`fileUploadStatus`} iconName="StatusCircleInner" style={{ color: `${this.fileStateColor(item.fileState)}` }}></Icon>
                                    <Icon iconName="Cancel" onClick={() => this.removeFile(item.file.name)}></Icon>
                                </Label>);
                            })
                    }
                </div>
                <div className={`spoFiles`}>
                    {
                        this.state.spoFiles.map(item => {
                            return (<Label key={item} title={this.fileStateText(IFileState.UPLOADED)} className={`fileName`}>{item}
                                <Icon className={`fileUploadStatus`} iconName="StatusCircleInner" style={{ color: `${this.fileStateColor(IFileState.UPLOADED)}` }}></Icon>
                                <Icon iconName="Cancel" onClick={() => this.deleteSpoFiles(item)}></Icon>
                            </Label>);
                        })
                    }

                </div>
            </div>
        );
    }

    /**
     * Removes the clicked file from the collection
     * @param fileName 
     */
    private removeFile(fileName: string): void {

        let file: ISPFxFile[] = this.state.selectedFiles.filter(i => i.file.name != fileName);

        this.setState({ selectedFiles: file }, () => {
            if (this.props.onFileAdded) this.props.onFileAdded(this.state.selectedFiles);
        });
    }

    /**
     * Uploads the file to the SharePoint
     */
    private async uploadFilesToSP(): Promise<void> {

        if (this.props.spWeb && this.props.docLibInternalName && this.state.selectedFiles.length > 0) {

            let files: ISPFxFile[] = this.state.selectedFiles;
            let file: ISPFxFile = files.pop();

            try {
                let fileURl: string = `${this.props.spWeb.absoluteUrl}/_api/Web/GetFolderByServerRelativeUrl('${this.folderPath}')/files/Add(url='${file.file.name}', overwrite=true)`;

                let options: ISPHttpClientOptions = {
                    headers: { 'odata-version': '3.0' },
                    body: file.file
                };

                await this.props.spHttp.post(fileURl, SPHttpClient.configurations.v1, options);
                this.setState({ selectedFiles: files, spoFiles: [file.file.name, ...this.state.spoFiles] });
                this.uploadFilesToSP();
            }
            catch (error) {
                Log.error(LOG_SOURCE, new Error(`Error occured while uploading the file: ${file.file.name}`));
                Log.error(LOG_SOURCE, error);
                file.fileState = IFileState.ERROR;
            }
            finally {
                this.getFilesOnSpo();
            }

        }
    }

    /**
     * Deletes teh file from the SHarePoint Library
     * @param fileName 
     */
    private async deleteSpoFiles(fileName: string): Promise<void> {

        let cnfrm: boolean = confirm('Do you want to delete the uploaded file from library? This action cannot be undone.');
        try {
            if (cnfrm) {

                let fileURl: string = `${this.props.spWeb.absoluteUrl}/_api/Web/GetFileByServerRelativeUrl('${this.props.spWeb.serverRelativeUrl}/${this.folderPath}/${fileName}')`;

                let options: ISPHttpClientOptions = {
                    headers: {
                        'odata-version': '3.0',
                        "IF-MATCH": "*",
                        'X-HTTP-Method': 'DELETE'
                    }
                };

                // Delete  the file and then update the state of the available files
                await this.props.spHttp.post(fileURl, SPHttpClient.configurations.v1, options);
                this.getFilesOnSpo();
            }
        }
        catch (error) {
            Log.error(LOG_SOURCE, error);
        }
    }

    /**
     * Returns the files that are on the SharePoint Library
     */
    private async getFilesOnSpo(): Promise<void> {

        try {

            let url: string = `${this.props.spWeb.absoluteUrl}/_api/Web/GetFolderByServerRelativeUrl('${this.folderPath}')/files`;

            let options: ISPHttpClientOptions = {
                headers: { 'odata-version': '3.0' }
            };

            let response: SPHttpClientResponse = await this.props.spHttp.get(url, SPHttpClient.configurations.v1, options);
            let respJSON: any = await response.json();
            let existingFiles: string[] = [];

            if (response.ok) {
                (respJSON.value as any[]).forEach(i => {
                    existingFiles.push(i.Name);
                });

                this.setState({ spoFiles: existingFiles });
            }

        }
        catch (error) {
            Log.error(LOG_SOURCE, new Error(`Error occured while files from SPO from folder: ${this.props.folder}`));
            Log.error(LOG_SOURCE, error);
        }
    }
}

export { SPFxFileUploader };