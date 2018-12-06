import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { autobind } from "@uifabric/utilities/lib";
import { Rating, RatingSize } from "office-ui-fabric-react/lib/Rating";
import { Slider } from "office-ui-fabric-react/lib/Slider";
import { SpinButton } from "office-ui-fabric-react/lib/SpinButton";
import * as React from "react";


export interface INumberFilterProps {
  value: number;
  displayAs: string;
  min: number;
  max: number;
  step: number;
  decimalPlaces: number;
  suffix: string;
  showSliderValue: boolean;
  sliderVertical: boolean;
  sliderHeight: number;
  sliderAlign: string;
  largeStars: boolean;
  //starIcon: string; //currently broken in UI Fabric Rating Control
  labelShow: boolean;
  labelText: string;
  onChange: (value: number) => void;
  displayMode: DisplayMode;
  updateTitle: (title: string) => void;
  title: string;
}

export class NumberFilter extends React.Component<INumberFilterProps, {}> {
  public render(): React.ReactElement<INumberFilterProps> {
    return (
      <div>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateTitle}/>
        {this.props.displayAs === "slider" &&
          <div style={this.props.sliderVertical ? {height:`${this.props.sliderHeight}px`, display:'flex', justifyContent:this.props.sliderAlign} : undefined}>
            <Slider
              label={this.props.labelShow ? this.props.labelText : undefined}
              min={this.props.min}
              max={this.props.max}
              step={this.props.step}
              showValue={this.props.showSliderValue}
              value={this.props.value}
              vertical={this.props.sliderVertical}
              onChange={this.props.onChange}/>
          </div>
        }
        {this.props.displayAs === "field" &&
          <SpinButton
            label={this.props.labelShow ? this.props.labelText : undefined}
            value={this.formatValueString(this.props.value)}
            onValidate={this.onSpinValidate}
            onIncrement={this.onSpinIncrement}
            onDecrement={this.onSpinDecrement}/>
        }
        {this.props.displayAs === "rating" &&
          <Rating
            size={this.props.largeStars ? RatingSize.Large : RatingSize.Small}
            rating={this.props.value}
            max={this.props.max}
            onChanged={this.props.onChange}/>
        }
      </div>
    );
  }

  @autobind
  private onSpinValidate(rawValue: string): string {
    let numValue = this.extractNumValue(rawValue);
    return this.validateNumber(numValue);
  }

  private validateNumber(numValue: number): string {
    if(numValue > this.props.max) {
      numValue = this.props.max;
    }
    if(numValue < this.props.min) {
      numValue = this.props.min;
    }

    numValue = +numValue.toFixed(this.props.decimalPlaces);
    if(numValue !== this.props.value) {
      this.props.onChange(numValue);
    }

    return this.formatValueString(numValue);
  }

  @autobind
  private onSpinDecrement(rawValue: string) : string {
    let numValue = this.extractNumValue(rawValue);
    return this.validateNumber(numValue - this.props.step);
  }

  @autobind
  private onSpinIncrement(rawValue: string): string {
    let numValue = this.extractNumValue(rawValue);
    return this.validateNumber(numValue + this.props.step);
  }

  private extractNumValue(rawValue: string): number {
    let numValue: number;
    let baseValue: string = this.removeSuffix(rawValue);

    if(isNaN(+baseValue)){
      numValue = this.props.min;
    } else {
      numValue = +baseValue;
    }

    return numValue;
  }

  private hasSuffix(rawValue: string): boolean {
    if(!this.props.suffix) {
      return false;
    }

    let subString: string = rawValue.substr(rawValue.length - this.props.suffix.length);
    return subString === this.props.suffix;
  }

  private removeSuffix(rawValue: string): string {
    if(!this.hasSuffix(rawValue)) {
      return rawValue;
    }

    return rawValue.substr(0, rawValue.length - this.props.suffix.length);
  }

  private formatValueString(numValue: number): string {
		return this.addSuffix(numValue.toFixed(this.props.decimalPlaces));
	}

	private addSuffix(stringValue: string): string {
		if(!this.props.suffix){
			return stringValue;
		}

		return stringValue + this.props.suffix;
	}
}
