import { Version } from '@microsoft/sp-core-library';
import {
  IReadonlyTheme,
  ThemeChangedEventArgs,
  ThemeProvider
} from '@microsoft/sp-component-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './FeedbackSubmissionWebPart.module.scss';
import * as strings from 'FeedbackSubmissionWebPartStrings';
import { SPService } from './services/SPService';

interface IFeedbackData {
  q1: number;
  q2: string;
  q3: string;
  q4: string;
  q5: string;
}

export interface IFeedbackSubmissionWebPartProps {
  targetList: string;
  targetQuestion1Field: string;
  targetQuestion2Field: string;
  targetQuestion3Field: string;
  targetQuestion4Field: string;
  targetQuestion5Field: string;
  targetUserAckField: string;
  questionsList: string;
  question1Field: string;
  question2Field: string;
  question3Field: string;
  question4Field: string;
  question5Field: string;
  debugMode: boolean;
}

export default class FeedbackSubmissionWebPart extends BaseClientSideWebPart<IFeedbackSubmissionWebPartProps> {

  private _spService!: SPService;
  private _themeProvider!: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  private listName: string = "Download_Tracking";    // Target list
  private questionList: string = "Feedbackquestions";

  private question1Text: string = "Question 1";
  private question2Text: string = "Question 2";
  private question3Text: string = "Question 3";
  private question4Text: string = "Question 4";
  private question5Text: string = "Question 5";

  // Dropdown options
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _fieldOptions: IPropertyPaneDropdownOption[] = [];
  private _targetFieldOptions: IPropertyPaneDropdownOption[] = [];

  public onInit(): Promise<void> {
    this._spService = new SPService(this.context.spHttpClient);

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit();
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  // -------------------------
  // MAIN RENDER
  // -------------------------
  public render(): void {
    this._renderAsync().catch(err => console.error(err));
  }

  private async _renderAsync(): Promise<void> {
    // If properties are set, use them. Otherwise fall back to defaults (or empty)
    this.listName = this.properties.targetList || this.listName;
    this.questionList = this.properties.questionsList || this.questionList;

    await this._loadQuestions();

    const trackingId = this._getTrackingIdFromEncryptedUrl();

    // Theme Colors
    const { semanticColors, palette } = this._themeVariant || {};

    // Initial state
    let isAlreadySubmitted = false;
    let savedData: IFeedbackData = { q1: 0, q2: "", q3: "", q4: "", q5: "" };

    if (trackingId) {
      const result = await this._checkExistingFeedback(trackingId);
      isAlreadySubmitted = result.isAlreadySubmitted;
      savedData = result.savedData;
    }

    const htmlContent = this._getHtmlContent(trackingId, isAlreadySubmitted, savedData, semanticColors, palette);
    this.domElement.innerHTML = htmlContent;

    if (trackingId) {
      this._bindEvents();
      this._wireStarEvents();

      // If already submitted, update stars & disable interaction
      if (isAlreadySubmitted && savedData.q1 > 0) {
        this._updateStars(savedData.q1);
        const hiddenInput = this.domElement.querySelector('#q1rating') as HTMLInputElement;
        if (hiddenInput) hiddenInput.value = savedData.q1.toString();

        const starContainer = this.domElement.querySelector('#q1stars') as HTMLElement;
        if (starContainer) {
          starContainer.style.pointerEvents = 'none';
          starContainer.style.opacity = '0.7';
        }
      }
    }
  }

  private async _checkExistingFeedback(trackingId: string): Promise<{ isAlreadySubmitted: boolean, savedData: IFeedbackData }> {
    const userAckField = this.properties.targetUserAckField || "UserAck";
    const tQ1 = this.properties.targetQuestion1Field || "Question1";
    const tQ2 = this.properties.targetQuestion2Field || "Question2";
    const tQ3 = this.properties.targetQuestion3Field || "Question3";
    const tQ4 = this.properties.targetQuestion4Field || "Question4";
    const tQ5 = this.properties.targetQuestion5Field || "Question5";

    try {
      const item = await this._spService.getItem(
        this.context.pageContext.web.absoluteUrl,
        this.listName,
        trackingId,
        [userAckField, tQ1, tQ2, tQ3, tQ4, tQ5]
      );

      if (item?.[userAckField]) {
        return {
          isAlreadySubmitted: true,
          savedData: {
            q1: Number(item[tQ1]) || 0,
            q2: String(item[tQ2] || ""),
            q3: String(item[tQ3] || ""),
            q4: String(item[tQ4] || ""),
            q5: String(item[tQ5] || "")
          }
        };
      }
    } catch (err) {
      console.error("Error checking existing feedback:", err);
    }
    return { isAlreadySubmitted: false, savedData: { q1: 0, q2: "", q3: "", q4: "", q5: "" } };
  }

  private _getHtmlContent(trackingId: string | undefined, isAlreadySubmitted: boolean, savedData: IFeedbackData, semanticColors: IReadonlyTheme['semanticColors'] | undefined, palette: IReadonlyTheme['palette'] | undefined): string {
    const bodyText = semanticColors?.bodyText || "#000000";
    const bodyBackground = semanticColors?.bodyBackground || "#ffffff";
    const inputText = semanticColors?.inputText || "#000000";
    const inputBackground = semanticColors?.inputBackground || "#ffffff";
    const inputBorder = semanticColors?.inputBorder || '#666';
    const buttonBackground = semanticColors?.primaryButtonBackground || "#0078d4";
    const buttonText = semanticColors?.primaryButtonText || "#ffffff";
    const errorText = semanticColors?.errorText || "red";
    const bodySubtext = semanticColors?.bodySubtext || '#999';
    const themePrimary = palette?.themePrimary || '#f5b301';

    const userEmail = this.context.pageContext.user.email;
    const feedbackMsg = isAlreadySubmitted ? "✅ Feedback already submitted." : "";

    // Prepare state variables for template
    const disabledAttr = isAlreadySubmitted ? "disabled" : "";
    const btnBg = isAlreadySubmitted ? "#cccccc" : buttonBackground;
    const btnCursor = isAlreadySubmitted ? "default" : "pointer";

    // Pre-calculate selected states to avoid nested ternaries
    const q2YesSelected = savedData.q2 === "Yes" ? "selected" : "";
    const q2NoSelected = savedData.q2 === "No" ? "selected" : "";
    const q3HoursSelected = savedData.q3 === "Hours Saved" ? "selected" : "";
    const q3KnowledgeSelected = savedData.q3 === "Gained New Knowledge" ? "selected" : "";
    const q3MoneySelected = savedData.q3 === "$ Saved" ? "selected" : "";

    const savedQ4 = savedData.q4 || "";
    const savedQ5 = savedData.q5 || "";

    // Render logic
    const innerContent = trackingId ? `
            <p><strong>Logged in user:</strong> ${userEmail}</p>

            <!-- Q1: Rating (1-5 stars) -->
            <div style="margin-bottom:12px;">
              <label><strong>${this.escapeHtml(this.question1Text)}</strong> <span style="color:${errorText}">*</span></label><br/>
              <div id="q1stars" style="font-size:22px; cursor:pointer; user-select:none;">
                <span class="star" data-value="1">★</span>
                <span class="star" data-value="2">★</span>
                <span class="star" data-value="3">★</span>
                <span class="star" data-value="4">★</span>
                <span class="star" data-value="5">★</span>
              </div>
              <input type="hidden" id="q1rating" value="" />
            </div>

            <!-- Q2: Yes/No -->
            <div style="margin-bottom:12px;">
              <label><strong>${this.escapeHtml(this.question2Text)}</strong> <span style="color:${errorText}">*</span></label><br/>
              <select id="q2yesno" ${disabledAttr} style="padding:5px; width:150px; background-color:${inputBackground}; color:${inputText}; border: 1px solid ${inputBorder};">
                <option value="">--Select--</option>
                <option value="Yes" ${q2YesSelected}>Yes</option>
                <option value="No" ${q2NoSelected}>No</option>
              </select>
            </div>

            <!-- Q3: Impact category -->
            <div style="margin-bottom:12px;">
              <label><strong>${this.escapeHtml(this.question3Text)}</strong> <span style="color:${errorText}">*</span></label><br/>
              <select id="q3impact" ${disabledAttr} style="padding:5px; width:250px; background-color:${inputBackground}; color:${inputText}; border: 1px solid ${inputBorder};">
                <option value="">--Select Impact--</option>
                <option value="Hours Saved" ${q3HoursSelected}>Hours Saved</option>
                <option value="Gained New Knowledge" ${q3KnowledgeSelected}>Gained New Knowledge</option>
                <option value="$ Saved" ${q3MoneySelected}>$ Saved</option>
              </select>
            </div>

            <!-- Q4: Value to share (optional) -->
            <div style="margin-bottom:12px;">
              <label><strong>${this.escapeHtml(this.question4Text)}</strong></label><br/>
              <textarea id="q4text" rows="3" ${disabledAttr} style="width:100%; background-color:${inputBackground}; color:${inputText}; border: 1px solid ${inputBorder};">${savedQ4}</textarea>
            </div>

            <!-- Q5: How can we do better? (optional) -->
            <div style="margin-bottom:12px;">
              <label><strong>${this.escapeHtml(this.question5Text)}</strong></label><br/>
              <textarea id="q5text" rows="3" ${disabledAttr} style="width:100%; background-color:${inputBackground}; color:${inputText}; border: 1px solid ${inputBorder};">${savedQ5}</textarea>
            </div>

            <div style="margin-top:8px;">
              <button id="btnSubmit" ${disabledAttr} style="padding:10px 18px; background-color:${btnBg}; color:${buttonText}; border:none; cursor:${btnCursor}; border-radius:4px;">
                Submit
              </button>
            </div>
    `
      : `<div style="color:${errorText};">❌ Missing or invalid TrackingID in the URL</div>`;


    return `
      <div class="${styles.feedbackSubmission}" style="color:${bodyText}; background-color:${bodyBackground}; padding: 15px;">
        <h3>Feedback Form</h3>
        <div id="feedbackContainer">
          ${innerContent}
        </div>
        <div id="messageBox" style="margin-top:12px; font-weight:bold;">
          ${isAlreadySubmitted ? `<span style="color:green;">${feedbackMsg}</span>` : ""}
        </div>
        <style>
          .star { font-size: 26px; color: ${bodySubtext}; padding: 2px; }
          .star.selected { color: ${themePrimary}; }
        </style>
      </div>
    `;
  }

  // -------------------------
  // Load 5 question captions from Feedbackquestions list
  // -------------------------
  private async _loadQuestions(): Promise<void> {
    const qList = this.properties.questionsList || this.questionList;
    if (!qList) return;

    try {
      const q = await this._spService.getQuestions(this.context.pageContext.web.absoluteUrl, qList);
      if (q) {
        const f1 = this.properties.question1Field || "Question1";
        const f2 = this.properties.question2Field || "Question2";
        const f3 = this.properties.question3Field || "Question3";
        const f4 = this.properties.question4Field || "Question4";
        const f5 = this.properties.question5Field || "Question5";

        this.question1Text = q[f1] || this.question1Text;
        this.question2Text = q[f2] || this.question2Text;
        this.question3Text = q[f3] || this.question3Text;
        this.question4Text = q[f4] || this.question4Text;
        this.question5Text = q[f5] || this.question5Text;
      }
    } catch (error) {
      console.error("Error loading questions:", error);
    }
  }

  // -------------------------
  // Bind submit button event
  // -------------------------
  private _bindEvents(): void {
    const button = this.domElement.querySelector('#btnSubmit');
    if (button) {
      button.addEventListener('click', () => this._submitFeedback());
    }
  }

  // -------------------------
  // Setup star rating click/hover behavior
  // -------------------------
  private _wireStarEvents(): void {
    const stars = this.domElement.querySelectorAll('.star');
    const hiddenInput = this.domElement.querySelector('#q1rating') as HTMLInputElement;

    stars.forEach((el) => {
      el.addEventListener('click', (ev) => {
        const target = ev.currentTarget as HTMLElement;
        const value = target.dataset.value;
        if (!value) return;
        this._updateStars(Number(value));
        if (hiddenInput) hiddenInput.value = value;
      });

      el.addEventListener('mouseover', (ev) => {
        const target = ev.currentTarget as HTMLElement;
        const value = target.dataset.value;
        if (!value) return;
        this._updateStars(Number(value));
      });

      el.addEventListener('mouseout', () => {
        // restore to selected value
        const selected = hiddenInput ? Number(hiddenInput.value) || 0 : 0;
        this._updateStars(selected);
      });
    });

    // initialize stars to 0
    this._updateStars(0);
  }

  private _updateStars(n: number): void {
    const stars = this.domElement.querySelectorAll('.star');
    stars.forEach(s => {
      const el = s as HTMLElement;
      const v = Number(el.dataset.value);
      if (v <= n) el.classList.add('selected'); else el.classList.remove('selected');
    });
  }

  // -------------------------
  // Submit feedback — validates required, checks UserAck and updates item
  // -------------------------
  private async _submitFeedback(): Promise<void> {
    const messageBox = this.domElement.querySelector('#messageBox') as HTMLElement;
    messageBox.innerHTML = "";
    messageBox.style.color = "black";

    const trackingId = this._getTrackingIdFromEncryptedUrl();
    if (!trackingId) {
      messageBox.innerHTML = "❌ Missing or invalid TrackingID in the URL.";
      messageBox.style.color = "red";
      return;
    }

    // collect values
    const q1 = (this.domElement.querySelector('#q1rating') as HTMLInputElement).value; // rating 1-5
    const q2 = (this.domElement.querySelector('#q2yesno') as HTMLSelectElement).value;
    const q3 = (this.domElement.querySelector('#q3impact') as HTMLSelectElement).value;
    const q4 = (this.domElement.querySelector('#q4text') as HTMLTextAreaElement).value.trim();
    const q5 = (this.domElement.querySelector('#q5text') as HTMLTextAreaElement).value.trim();

    // validate mandatory questions 1-3
    if (!q1 || !q2 || !q3) {
      messageBox.innerHTML = "⚠ Please answer all mandatory questions (1–3).";
      messageBox.style.color = "red";
      return;
    }

    const targetList = this.properties.targetList || this.listName;
    const userAckField = this.properties.targetUserAckField || "UserAck";

    // Step 1: Fetch the target item by ID and check UserAck
    const item = await this._spService.getItem(this.context.pageContext.web.absoluteUrl, targetList, trackingId, [userAckField]);

    if (!item) {
      messageBox.innerHTML = "❌ Invalid Tracking ID. Item not found.";
      messageBox.style.color = "red";
      return;
    }

    if (item[userAckField]) {
      messageBox.innerHTML = "⚠ Feedback already provided for this document.";
      messageBox.style.color = "orange";
      return;
    }



    // Field internal names for target list
    const itemType = item.__metadata?.type || "SP.ListItem";

    const tQ1 = this.properties.targetQuestion1Field || "Question1";
    const tQ2 = this.properties.targetQuestion2Field || "Question2";
    const tQ3 = this.properties.targetQuestion3Field || "Question3";
    const tQ4 = this.properties.targetQuestion4Field || "Question4";
    const tQ5 = this.properties.targetQuestion5Field || "Question5";

    const values: Record<string, unknown> = {
      // Save values to the mapped list fields
      [tQ1]: q1,
      [tQ2]: q2,
      [tQ3]: q3,
      [tQ4]: q4,
      [tQ5]: q5,
      [userAckField]: true
    };



    const success = await this._spService.submitFeedback(this.context.pageContext.web.absoluteUrl, targetList, trackingId, values, itemType);

    if (success) {
      messageBox.innerHTML = "✅ Feedback submitted successfully!";
      messageBox.style.color = "green";

      // disable button to prevent re-submit
      const btn = this.domElement.querySelector('#btnSubmit') as HTMLButtonElement;
      if (btn) btn.disabled = true;
    } else {
      messageBox.innerHTML = `⚠ Failed to submit feedback.`;
      messageBox.style.color = "red";
    }
  }

  // -------------------------
  // Decrypt encrypted value directly after "?"
  // expects URL like: ?eyJUcmFja2luZ0lEIjoxfQ%3D%3D
  // -------------------------
  private _getTrackingIdFromEncryptedUrl(): string | undefined {
    try {
      // DEBUG MODE: Return exact param from URL
      if (this.properties.debugMode) {
        const params = new URLSearchParams(globalThis.location.search);
        return params.get("TrackingID") || undefined;
      }

      // Normal Mode: Encrypted
      // get the part after "?"
      let encrypted = globalThis.location.search ? globalThis.location.search.substring(1).trim() : "";
      if (!encrypted) return undefined;

      // URL decode to convert %3D%3D -> ==
      encrypted = decodeURIComponent(encrypted);

      // base64 decode
      const base64Decoded = atob(encrypted);

      // parse JSON
      const parsed = JSON.parse(base64Decoded);

      return parsed.TrackingID ? parsed.TrackingID.toString() : undefined;
    } catch (err: unknown) {
      console.error("❌ Failed to extract TrackingID:", err);
      return undefined;
    }
  }

  // small helper to escape any caption text inserted into HTML
  // small helper to escape any caption text inserted into HTML
  private escapeHtml(input: string): string {
    if (!input) return "";
    return input.replace(/[&<>"']/g, (match: string) => {
      switch (match) {
        case '&': return '&amp;';
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '"': return '&quot;';
        case "'": return '&#039;';
        default: return match;
      }
    });
  }

  // -------------------------
  // PROPERTY PANE CONFIGURATION
  // -------------------------
  protected onPropertyPaneConfigurationStart(): void {
    console.log('FeedbackSubmission: onPropertyPaneConfigurationStart');
    // If not yet initialized, load lists.
    // NOTE: Property pane might open before data is ready.
    // We should trigger a refresh once data is loaded.
    this._loadLists().then(() => {
      console.log('FeedbackSubmission: _loadLists completed');
      // Check if we need to load fields too
      if (this.properties.questionsList) {
        console.log('FeedbackSubmission: Loading fields for questionsList', this.properties.questionsList);
        this._loadFields(this.properties.questionsList, false).catch(e => console.error(e));
      }
      if (this.properties.targetList) {
        console.log('FeedbackSubmission: Loading fields for targetList', this.properties.targetList);
        this._loadFields(this.properties.targetList, true).catch(e => console.error(e));
      }
    }).catch(err => {
      console.error('FeedbackSubmission: _loadLists failed', err);
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath === 'questionsList' && newValue) {
      this.properties.questionsList = newValue as string;
      this._fieldOptions = [];
      this.context.propertyPane.refresh();

      this._loadFields(newValue as string).catch(err => console.error(err));
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  private async _loadLists(): Promise<void> {
    console.log('FeedbackSubmission: _loadLists called');
    try {
      this._listOptions = await this._spService.getLists(this.context.pageContext.web.absoluteUrl);
      console.log('FeedbackSubmission: Lists loaded', this._listOptions.length);
      this.context.propertyPane.refresh();
    } catch (e) {
      console.error('FeedbackSubmission: Error in _loadLists', e);
    }
  }

  private async _loadFields(listTitle: string, isTarget: boolean = false): Promise<void> {
    console.log(`FeedbackSubmission: _loadFields called for '${listTitle}' (isTarget=${isTarget})`);
    try {
      const fields = await this._spService.getFields(this.context.pageContext.web.absoluteUrl, listTitle);
      console.log(`FeedbackSubmission: Fields loaded for '${listTitle}': ${fields.length}`);
      if (isTarget) {
        this._targetFieldOptions = fields;
      } else {
        this._fieldOptions = fields;
      }
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error(`FeedbackSubmission: Error loading fields for '${listTitle}'`, error);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "General Settings",
              groupFields: [
                PropertyPaneDropdown('targetList', {
                  label: "Target List (Feedback Submission)",
                  options: this._listOptions,
                  selectedKey: this.properties.targetList,
                  disabled: this._listOptions.length === 0
                }),
                PropertyPaneDropdown('targetQuestion1Field', {
                  label: "Target Question 1 Field",
                  options: this._targetFieldOptions,
                  disabled: !this.properties.targetList || this._targetFieldOptions.length === 0
                }),
                PropertyPaneDropdown('targetQuestion2Field', {
                  label: "Target Question 2 Field",
                  options: this._targetFieldOptions,
                  disabled: !this.properties.targetList || this._targetFieldOptions.length === 0
                }),
                PropertyPaneDropdown('targetQuestion3Field', {
                  label: "Target Question 3 Field",
                  options: this._targetFieldOptions,
                  disabled: !this.properties.targetList || this._targetFieldOptions.length === 0
                }),
                PropertyPaneDropdown('targetQuestion4Field', {
                  label: "Target Question 4 Field",
                  options: this._targetFieldOptions,
                  disabled: !this.properties.targetList || this._targetFieldOptions.length === 0
                }),
                PropertyPaneDropdown('targetQuestion5Field', {
                  label: "Target Question 5 Field",
                  options: this._targetFieldOptions,
                  disabled: !this.properties.targetList || this._targetFieldOptions.length === 0
                }),
                PropertyPaneDropdown('targetUserAckField', {
                  label: "Target UserAck Field",
                  options: this._targetFieldOptions,
                  disabled: !this.properties.targetList || this._targetFieldOptions.length === 0
                }),
                PropertyPaneToggle('debugMode', {
                  label: "Debug Mode",
                  onText: "Yes",
                  offText: "No",
                  checked: this.properties.debugMode
                })
              ]
            },
            {
              groupName: "Question Configuration",
              groupFields: [
                PropertyPaneDropdown('questionsList', {
                  label: "Questions Source List",
                  options: this._listOptions,
                  selectedKey: this.properties.questionsList
                }),
                PropertyPaneDropdown('question1Field', {
                  label: "Question 1 Field",
                  options: this._fieldOptions,
                  disabled: !this.properties.questionsList || this._fieldOptions.length === 0
                }),
                PropertyPaneDropdown('question2Field', {
                  label: "Question 2 Field",
                  options: this._fieldOptions,
                  disabled: !this.properties.questionsList || this._fieldOptions.length === 0
                }),
                PropertyPaneDropdown('question3Field', {
                  label: "Question 3 Field",
                  options: this._fieldOptions,
                  disabled: !this.properties.questionsList || this._fieldOptions.length === 0
                }),
                PropertyPaneDropdown('question4Field', {
                  label: "Question 4 Field",
                  options: this._fieldOptions,
                  disabled: !this.properties.questionsList || this._fieldOptions.length === 0
                }),
                PropertyPaneDropdown('question5Field', {
                  label: "Question 5 Field",
                  options: this._fieldOptions,
                  disabled: !this.properties.questionsList || this._fieldOptions.length === 0
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}