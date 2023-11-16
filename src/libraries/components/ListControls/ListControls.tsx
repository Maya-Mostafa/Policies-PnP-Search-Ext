import * as React from 'react';
import {CommandBarButton, Toggle, DialogType} from 'office-ui-fabric-react';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { initializeIcons } from '@uifabric/icons';
import {IListControlsProps} from './IListControlsProps';
import styles from '../../component.module.scss';

export default function ListControls (props: IListControlsProps) {
  
  initializeIcons();

  return (
    <div className={styles.listControls}>
            
      <CommandBarButton iconProps={{iconName: 'CloudUpload'}} text="Upload Document" onClick={props.uploadDocumentHandler} />
      <CommandBarButton iconProps={{ iconName: 'Documentation' }} text="View All" onClick={props.viewAllHandler} />

      <IFrameDialog 
          url={props.iFrameUrl}
          width={'70%'}
          height={'90%'}
          hidden={!props.iFrameVisible}
          onDismiss={() => props.setIFrameVisible(false)}
          allowFullScreen = {true}
          dialogContentProps={{
            type: DialogType.close,
            showCloseButton: true
          }}
      />

    </div>
  );

}