/* eslint-disable @typescript-eslint/no-explicit-any */
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import styles from '../Employeedetails.module.scss';
import { IFieldInfo } from '@pnp/sp/fields';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import { getSP } from '../../../../pnpjs-config'; // Import your PnPjs configuration
import DynamicForm from '../DynamicForm/DynamicForm'; // Import the DynamicForm component
import EmployeeService from '../../../../service/EmployeeService';

export interface IFormContainerState {
    fields: IFieldInfo[];
    formData: { [key: string]: any }; // Store form data dynamically
}

export interface IFormContainerProps {
    context: FormCustomizerContext;
    displayMode: FormDisplayMode;
    onSave: () => void;
    onClose: () => void;
}

export default class FormContainer extends React.Component<IFormContainerProps, IFormContainerState> {
    private _sp: SPFI;

    constructor(props: IFormContainerProps) {
        super(props);
        this.state = {
            fields: [],
            formData: {},
        };
        this._sp = getSP(this.props.context); // Initialize _sp with the SPFI instance
    }

    public componentDidMount(): void {
        this._getFields()
            .then((filteredFields) => {
                this.setState({ fields: filteredFields });
            })
            .catch((error) => {
                console.error("Error fetching fields:", error);
            });
    }

    private async _getFields(): Promise<IFieldInfo[]> {
        const targetFields = ["Title","EmployeeFirstName", "Qualification", "StartDate1", "ABS", "Manager", "ManagerComments"];
        try {
            const allFields = await this._sp.web.lists.getById(this.props.context.list.guid.toString()).fields();
            const filteredFields = allFields.filter((field) =>
                targetFields.includes(field.InternalName)
            );
            return filteredFields;
        } catch (error) {
            console.error("Error fetching fields:", error);
            throw error;
        }
    }

    private _getValueForUpdate(fieldInfo: IFieldInfo, value: any): any {
        if (!value) return null;

        if (fieldInfo.TypeAsString === "TaxonomyFieldType") {
            // Single-value taxonomy
            if (Array.isArray(value) && value.length > 0) {
                return {
                    Label: value[0].Label,
                    TermGuid: value[0].TermGuid,
                    WssId: value[0].WssId || -1, // Default WssId to -1 for new items
                };
            }
            return null;
        }

        if (fieldInfo.TypeAsString === "TaxonomyFieldTypeMulti") {
            // Multi-value taxonomy
            if (Array.isArray(value) && value.length > 0) {
                return {
                    results: value.map((term) => ({
                        Label: term.Label,
                        TermGuid: term.TermGuid,
                        WssId: term.WssId || -1, // Default WssId to -1 for new items
                    })),
                };
            }
            return null;
        }
        if (fieldInfo.TypeAsString === "User") {
            if (Array.isArray(value) && value.length > 0) {
                return value[0].Id; // Take the first user from the array
            }
            return typeof value === "object" && value.Id ? value.Id : value;
        }

        if (fieldInfo.TypeAsString === "UserMulti") {
            return Array.isArray(value)
                ? {
                    results: value.map((user) =>
                        typeof user === "object" ? user.Id : user
                    ),
                }
                : null;
        }

        return value; // Default for other field types
    }

    private handleSave = (formData: any): void => {
        const requiredFields = this.state.fields.filter((field) => field.Required);
        const missingFields = requiredFields.filter((field) => !formData[field.InternalName]);

        if (missingFields.length > 0) {
            alert(`Please fill in all required fields: ${missingFields.map((field) => field.Title).join(", ")}`);
            return;
        }

        const createObject: any = {};
        Object.entries(formData).forEach(([key, value]: [string, any]) => {
            const field = this.state.fields.find((f) => f.InternalName === key);
            if (!field) return;

            const formattedValue = this._getValueForUpdate(field, value);

            if (formattedValue !== null) {
                // Rename user field keys to "FieldNameId"
                if (field.TypeAsString === "User") {
                    createObject[`${key}Id`] = formattedValue;
                } else if (field.TypeAsString === "UserMulti") {
                    createObject[`${key}Id`] = formattedValue;
                } else {
                    createObject[key] = formattedValue;
                }
            }
        });
        console.log("Formatted createObject:", createObject); // Debugging line
        // const createObject1 = {
        //     Title: "Test 2",
        //     Qualification: "Degree",
        //     StartDate1: new Date("2025-04-07"),
        //     ABS: {
        //         Label: "ABS1",
        //         TermGuid: "88e4fa67-8856-449b-ac7b-0e08a4918286",
        //         WssId: -1
        //     },
        //     ManagerId: 12
        // };

        EmployeeService.createListItem("EmployeeDetails", createObject)
            .then((res) => {
                console.log("Item created successfully:", res);
                alert("Item created successfully");
                this.props.onSave();
            })
            .catch((error) => {
                alert("Error creating item");
                console.error("Error creating item:", error);
            });
    };

    public render(): React.ReactElement<{}> {
        const { fields } = this.state;

        return (
            <div className={styles.formWrapper}>
                <div className={styles.formOuterContainer}>
                    <div className={styles.formMidContainer}>
                        {fields.length > 0 ? (
                            <DynamicForm
                                fields={fields}
                                onSave={this.handleSave}
                                onClose={this.props.onClose}
                                context={this.props.context}
                            />
                        ) : (
                            <div>Loading fields...</div>
                        )}
                    </div>
                </div>
            </div>
        );
    }
}
