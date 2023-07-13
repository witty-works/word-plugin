import { Pipe, PipeTransform } from '@angular/core';
import TextUtils from "../utils/text.utils";

@Pipe({
  name: 'highlight'
})
export class HighlightPipe implements PipeTransform {

  transform(value: any, args: string): unknown {
    if(!value || !args) return value;
    const re = new RegExp(`(?<![äöüÄÖÜàéèòìÀÉÈÒÌ\\w])(${TextUtils.escapeRegExp(args)})(?![äöüÄÖÜàéèòìÀÉÈÒÌ\\w])`, 'gm');
    const color = this.getExplanationColor('null'); //TODO: need to figure out how to get graviy here 
    value = value.replace(re, `<span class="highlighted-text highlighted-text--${color}">$1</span>`);
    return value;
  }

  getExplanationColor(gravity: string): string {
    let gravityNum = parseFloat(gravity);
    if (!gravityNum) return 'green';
    else if (gravityNum < 1.5) return 'red';
    else if (gravityNum > 2.5) return 'yellow';
    else return 'orange';
  }
}
