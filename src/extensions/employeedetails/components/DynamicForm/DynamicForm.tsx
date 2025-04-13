/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";

import { TextField } from "@fluentui/react/lib/TextField";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { DatePicker } from "@fluentui/react/lib/DatePicker";
 import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ModernTaxonomyPicker } from "@pnp/spfx-controls-react/lib/ModernTaxonomyPicker";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { TooltipHost, TooltipDelay } from "@fluentui/react/lib/Tooltip";
import { IFieldInfo } from "@pnp/sp/fields";
import styles from "../Employeedetails.module.scss";
import { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IPersonaProps } from "@fluentui/react";
export interface IDynamicFormProps {
    fields: IFieldInfo[]; // Fields retrieved from SharePoint
    onSave: (formData: any) => void; // Callback to handle form submission
    onClose: () => void; // Callback to handle form cancellation
    context: any; // SPFx context for PeoplePicker and ModernTaxonomyPicker
}

export interface IDynamicFormState {
    formData: { [key: string]: any }; // Stores the values of the form fields
}

export default class DynamicForm extends React.Component<IDynamicFormProps, IDynamicFormState> {
    protected _peoplePickerContext: IPeoplePickerContext;

    constructor(props: IDynamicFormProps) {
        super(props);
        this.state = {
            formData: {}, // Initialize form data as an empty object
        };
    }

    private handleInputChange = (fieldInternalName: string, value: any): void => {
        this.setState((prevState) => ({
            formData: {
                ...prevState.formData,
                [fieldInternalName]: value,
            },
        }));
    };

    private renderField = (field: IFieldInfo): React.ReactElement => {
        const { formData } = this.state;

        switch (field.TypeAsString) {
            case "Text": // Single line of text
                return (
                    <TextField
                        key={field.InternalName}
                        label={field.Title}
                        required={field.Required}
                        value={formData[field.InternalName] || ""}
                        onChange={(e, newValue) => this.handleInputChange(field.InternalName, newValue || "")}
                    />
                );

            case "Choice": { // Choice field
                const choiceOptions: IDropdownOption[] = (field.Choices || []).map((choice) => ({
                    key: choice,
                    text: choice,
                }));
                return (
                    <Dropdown
                        key={field.InternalName}
                        label={field.Title}
                        required={field.Required}
                        options={choiceOptions}
                        selectedKey={formData[field.InternalName]}
                        onChange={(e, option) => this.handleInputChange(field.InternalName, option?.key || "")}
                    />
                );
            }

            case "DateTime": // Date and Time field
                return (
                    <DatePicker
                        key={field.InternalName}
                        label={field.Title}
                        isRequired={field.Required}
                        value={formData[field.InternalName] ? new Date(formData[field.InternalName]) : undefined}
                        onSelectDate={(date) => this.handleInputChange(field.InternalName, date)}
                    />
                );

            case "User": // Single-select Person or Group field
            case "UserMulti": // Multi-select Person or Group field
                return (
                    <PeoplePicker
                        key={field.InternalName}
                        context={this.props.context}
                        titleText={field.Title}
                        personSelectionLimit={field.TypeAsString === "User" ? 1 : 20}
                        required={field.Required}
                        showtooltip={true}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]} // Include both users and groups
                        resolveDelay={500} // Add delay for better performance
                        suggestionsLimit={10}
                        webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
                        defaultSelectedUsers={
                            formData[field.InternalName]?.map((user: any) => user.Email) || []
                        } // Prepopulate with existing values
                        onChange={(items) =>
                            this._onPeoplePickerChange(
                                field.InternalName,
                                items.map((item) => ({
                                    ...item,
                                    loginName: item.id || "",
                                    email: item.secondaryText || "",
                                }))
                            ) // Map items to include loginName and email
                        }
                        ensureUser // Ensure users are resolved correctly
                    />
                );

            case "TaxonomyFieldType": // Managed Metadata field
            case "TaxonomyFieldTypeMulti": // Multi-select Managed Metadata field
                return (
                    <ModernTaxonomyPicker
                        key={field.InternalName}
                        allowMultipleSelections={field.TypeAsString === "TaxonomyFieldTypeMulti"}
                        termSetId={(field as IFieldInfo & { TermSetId: string }).TermSetId}
                        panelTitle={`Select ${field.Title}`}
                        label={field.Title}
                        context={this.props.context}
                        onChange={(terms) =>
                            this.handleInputChange(
                                field.InternalName,
                                (terms ?? []).map((term) => ({
                                    Label: term.labels?.[0]?.name || "", // Extract the first label name or fallback to an empty string
                                    TermGuid: term.id,
                                    WssId: -1, // WssId is typically set to -1 for new items
                                }))
                            )
                        }
                        required={field.Required}
                        initialValues={formData[field.InternalName] || []}
                        disabled={field.ReadOnlyField}
                        termPickerProps={{
                            itemLimit: field.TypeAsString === "TaxonomyFieldType" ? 1 : undefined,
                        }}
                    />
                );

            case "Note": // Multiple lines of text
                return (
                    <TextField
                        key={field.InternalName}
                        label={field.Title}
                        required={field.Required}
                        multiline
                        value={formData[field.InternalName] || ""}
                        onChange={(e, newValue) => this.handleInputChange(field.InternalName, newValue || "")}
                    />
                );

            default:
                return (
                    <TooltipHost
                        key={field.InternalName}
                        content={`Field type "${field.TypeAsString}" is not supported.`}
                        delay={TooltipDelay.long}
                    >
                        <TextField
                            label={field.Title}
                            disabled
                            placeholder={`Unsupported field type: ${field.TypeAsString}`}
                        />
                    </TooltipHost>
                );
        }
    };

    private _onPeoplePickerChange(fieldInternalName: string, items: (IPersonaProps & { loginName: string; email: string })[]): void {
        const newValue = items.map((item) => ({
            Email: item.secondaryText,
            Id: +(item.id || 0),
            Title: item.text,
            LoginName: item.loginName,
        }));

        this.setState((prevState) => ({
            formData: {
                ...prevState.formData,
                [fieldInternalName]: newValue,
            },
        }));
    }

    private handleSave = (): void => {
        this.props.onSave(this.state.formData); // Pass form data to the parent component
    };

    public render(): React.ReactElement {
        const { fields, onClose } = this.props;

        return (
            <div className={styles.formWrapper}>
                <div className={`${styles.formOuterContainer}`}>
                    <div className={`${styles.formMidContainer}`}>
                        {fields.map((field) => (
                            <div className={styles.fieldWrapper} key={field.InternalName}>
                                {this.renderField(field)}
                            </div>
                        ))}
                    </div>
                    <div className={styles.formFooter}>
                        <DefaultButton text="Cancel" onClick={onClose} />
                        <PrimaryButton text="Save" onClick={this.handleSave} />
                    </div>
                </div>
            </div>
        );
    }
}
