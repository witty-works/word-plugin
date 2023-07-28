import { Pipe, PipeTransform } from '@angular/core';
import TextUtils from "../utils/text.utils";

@Pipe({
  name: 'highlight'
})
export class HighlightPipe implements PipeTransform {

  transform(value: any, error: any): unknown {
    if(!value || !error.word) return value;
    const re = new RegExp(`(?<![äöüÄÖÜàéèòìÀÉÈÒÌ\\w])(${TextUtils.escapeRegExp(error.word)})(?![äöüÄÖÜàéèòìÀÉÈÒÌ\\w])`, 'gm');
    const color = this.getExplanationColor(error.details.gravity);
    value = value.replace(re, `<span class="highlighted-text highlighted-text--${color}">$1</span>`);
    return value;
  }

  getExplanationColor(gravity: number): string {
    if (!gravity) return 'green';
    else if (gravity < 1.5) return 'red';
    else if (gravity > 2.5) return 'yellow';
    else return 'orange';
  }
}
