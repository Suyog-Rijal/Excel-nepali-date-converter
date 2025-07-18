import React from "react";
import {
    Dropdown,
    Option,
    Label,
    makeStyles,
    tokens,
    DropdownProps
} from "@fluentui/react-components";

export interface CustomDropdownProps {
    id: string;
    label?: string;
    placeholder?: string;
    value?: string | null;
    disabled?: boolean;
    options?: { key: string; value: string }[];
    onChange?: (value: string) => void;
}

const useStyles = makeStyles({
    container: {
        display: "flex",
        flexDirection: "column",
        rowGap: "0.15rem",
        width: "100%",
        maxWidth: "250px"
    },
    label: {
        fontSize: "0.75rem",
        fontWeight: 400,
        color: tokens.colorNeutralForeground3
    },
    dropdown: {
        width: "100%",
        backgroundColor: tokens.colorNeutralBackground1
    }
});

const CustomDropdown = ({
                            id,
                            label,
                            placeholder,
                            value,
                            options,
                            onChange,
                            disabled = false
                        }: CustomDropdownProps) => {
    const styles = useStyles();

    const handleOptionSelect: DropdownProps["onOptionSelect"] = (_, data) => {
        if (onChange) {
            onChange(data.optionValue as string);
        }
    };

    const selectedOptionLabel =
        options?.find((opt) => opt.key === value)?.value ?? "";

    return (
        <div className={styles.container}>
            {label && (
                <Label htmlFor={id} className={styles.label}>
                    {label}
                </Label>
            )}
            <Dropdown
                disabled={disabled}
                id={id}
                placeholder={placeholder}
                className={styles.dropdown}
                value={selectedOptionLabel}
                onOptionSelect={handleOptionSelect}
            >
                {options?.map(option => (
                    <Option key={option.key} value={option.key}>
                        {option.value}
                    </Option>
                ))}
            </Dropdown>
        </div>
    );
};

export default CustomDropdown;
