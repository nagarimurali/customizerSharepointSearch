import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { TextField, PrimaryButton, DefaultButton, Stack } from '@fluentui/react';

export interface IDmsTranslatedocSetProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

export interface IDmsTranslatedocSetState {
  formData: {
    lastName: string;
    firstName: string;
    email: string;
    company: string;
    jobTitle: string;
    businessPhone: string;
    homePhone: string;
    mobileNumber: string;
    faxNumber: string;
    address: string;
    city: string;
    state: string;
    zip: string;
    country: string;
    webPage: string;
    comments: string;
  };
}

const LOG_SOURCE: string = 'DmsTranslatedocSet';

export default class DmsTranslatedocSet extends React.Component<IDmsTranslatedocSetProps, IDmsTranslatedocSetState> {
  constructor(props: IDmsTranslatedocSetProps) {
    super(props);
    this.state = {
      formData: {
        lastName: '',
        firstName: '',
        email: '',
        company: '',
        jobTitle: '',
        businessPhone: '',
        homePhone: '',
        mobileNumber: '',
        faxNumber: '',
        address: '',
        city: '',
        state: '',
        zip: '',
        country: '',
        webPage: '',
        comments: '',
      },
    };
  }

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: DmsTranslatedocSet mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: DmsTranslatedocSet unmounted');
  }

  private handleInputChange = (field: string, value: string): void => {
    this.setState((prevState) => ({
      formData: {
        ...prevState.formData,
        [field]: value,
      },
    }));
  };

  private handleSave = async (): Promise<void> => {
    const { formData } = this.state;
    const { context } = this.props;

    // Validate required fields
    if (!formData.lastName || !formData.firstName) {
      Log.error(LOG_SOURCE, new Error('Validation failed: Required fields are missing.'));
      alert('Please fill in all required fields: Last Name and First Name.');
      return;
    }

    try {
      const listTitle = context.list.title; // Ensure this matches the actual list name
      const requestUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;

      // Ensure Title field is populated and include ContentTypeId
      const payload = {
        Title: formData.lastName || "Untitled", // Default to "Untitled" if lastName is empty
        ContentTypeId: "0x0106007F78E7573CF49643805D371402DCD537", // Default content type ID for "Item" (adjust if using a custom content type)
        FirstName: formData.firstName,
        Email: formData.email || null, // Use null for empty optional fields
        Company: formData.company || null,
        JobTitle: formData.jobTitle || null,
        BusinessPhone: formData.businessPhone || null,
        HomePhone: formData.homePhone || null,
        MobileNumber: formData.mobileNumber || null,
        FaxNumber: formData.faxNumber || null,
        Address: formData.address || null,
        City: formData.city || null,
        State: formData.state || null,
        ZIP: formData.zip || null,
        Country: formData.country || null,
        WebPage: formData.webPage || null,
        Comments: formData.comments || null,
      };

      const response = await context.spHttpClient.post(
        requestUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': '',
          },
          body: JSON.stringify(payload),
        }
      );

      if (response.ok) {
        Log.info(LOG_SOURCE, 'Form data saved successfully.');
        this.props.onSave();
      } else {
        const error = await response.json();
        Log.error(LOG_SOURCE, new Error(`Error saving data: ${error.error.message}`));
        console.error('Error details:', error); // Log error details for debugging
      }
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Error saving data: ${error.message}`));
      console.error('Error details:', error); // Log error details for debugging
    }
  };

  public render(): React.ReactElement<{}> {
    const { displayMode, onClose } = this.props;
    const { formData } = this.state;

    const isReadOnly = displayMode === FormDisplayMode.Display;

    return (
      <div>
        <Stack tokens={{ childrenGap: 10 }}>
          <TextField
            label="Last Name"
            value={formData.lastName}
            onChange={(e, newValue) => this.handleInputChange('lastName', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="First Name"
            value={formData.firstName}
            onChange={(e, newValue) => this.handleInputChange('firstName', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="E-Mail"
            value={formData.email}
            onChange={(e, newValue) => this.handleInputChange('email', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Company"
            value={formData.company}
            onChange={(e, newValue) => this.handleInputChange('company', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Job Title"
            value={formData.jobTitle}
            onChange={(e, newValue) => this.handleInputChange('jobTitle', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Business Phone"
            value={formData.businessPhone}
            onChange={(e, newValue) => this.handleInputChange('businessPhone', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Home Phone"
            value={formData.homePhone}
            onChange={(e, newValue) => this.handleInputChange('homePhone', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Mobile Number"
            value={formData.mobileNumber}
            onChange={(e, newValue) => this.handleInputChange('mobileNumber', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Fax Number"
            value={formData.faxNumber}
            onChange={(e, newValue) => this.handleInputChange('faxNumber', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Address"
            value={formData.address}
            onChange={(e, newValue) => this.handleInputChange('address', newValue || '')}
            readOnly={isReadOnly}
            multiline
          />
          <TextField
            label="City"
            value={formData.city}
            onChange={(e, newValue) => this.handleInputChange('city', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="State/Province"
            value={formData.state}
            onChange={(e, newValue) => this.handleInputChange('state', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="ZIP/Postal Code"
            value={formData.zip}
            onChange={(e, newValue) => this.handleInputChange('zip', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Country/Region"
            value={formData.country}
            onChange={(e, newValue) => this.handleInputChange('country', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Web Page"
            value={formData.webPage}
            onChange={(e, newValue) => this.handleInputChange('webPage', newValue || '')}
            readOnly={isReadOnly}
          />
          <TextField
            label="Comments"
            value={formData.comments}
            onChange={(e, newValue) => this.handleInputChange('comments', newValue || '')}
            readOnly={isReadOnly}
            multiline
          />
        </Stack>

        <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
          {!isReadOnly && (
            <PrimaryButton text="Save" onClick={this.handleSave} />
          )}
          <DefaultButton text="Close" onClick={onClose} />
        </Stack>
      </div>
    );
  }
}
