import * as React from 'react';

export const displayTemplateCallout = (): JSX.Element => {
    return (
        <span>
            Add additional text to the value by using <span className="ms-fontColor-themePrimary ms-fontWeight-semibold">{"[VALUE]"}</span> wherever you want the dynamic value. For instance:<br /><br />
            <span className="ms-fontWeight-semibold">Welcome to <span className="ms-fontColor-themePrimary">{"[VALUE]"}</span>!</span>
        </span>
    );
};
