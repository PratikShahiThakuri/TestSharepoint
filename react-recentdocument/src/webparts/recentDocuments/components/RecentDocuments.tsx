import * as React from 'react';
import styles from './RecentDocuments.module.scss';
import { IRecentDocumentsProps } from './IRecentDocumentsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IDocument } from '../../../documentsServices';
import { IRecentDocumentProps } from '../IRecentDocumentProps';
import {
  DocumentCard,
  DocumentCardType,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity
 } from 'office-ui-fabric-react';
export default class RecentDocuments extends React.Component<IRecentDocumentProps, {}> {
  public render(): React.ReactElement<IRecentDocumentsProps> {
    const documents: JSX.Element[] = this.props.documents.map((document: IDocument, index: number, array: IDocument[]): JSX.Element => {
      return (
        <DocumentCard type={DocumentCardType.compact} onClickHref={document.url} accentColor='#ce4b1f' key={index}>
          <DocumentCardPreview previewImages={[{
            name: document.title,
            url: document.url,
            previewImageSrc: document.imageUrl,
            iconSrc: document.iconUrl,
            width: 144
          }]} />
          <div className='ms-DocumentCard-details'>
            <DocumentCardTitle
              title={document.title}
              shouldTruncate={true} />
            <DocumentCardActivity
              activity={document.activity.title}
              people={
                [
                  { name: document.activity.actorName, profileImageSrc: document.activity.actorImageUrl }
                ]
              }
              />
          </div>
        </DocumentCard>
      );
    });
    return (
      <div className={styles.recentDocuments}>
        {documents}
      </div>
    );
  }
}
