# PGN chatbot QA

PGN Sawala executes prepared chatbot scenarios and records responses for human evaluation. These terms distinguish the test content, execution context, and kinds of result.

## Language

### Test content

**Scenario**: A test case identified by a Test Case ID, containing one or more ordered turns that share conversation context.
_Avoid_: Row or turn when referring to the whole test case.

**Turn**: One prepared user input and the response collected for it within a scenario. A response can contain several bot messages.
_Avoid_: Treating each bot message as a separate turn.

**Source workbook**: The prepared test inputs, expected responses, and reference material from which an executed workbook is created. Execution preserves this workbook.

**Executed workbook**: The working results copy containing captured responses, evaluation statuses, transcripts, and run history. It can accumulate results from multiple runs.

### Execution context

**Run**: One execution of a selected set of scenarios, identified by a Run ID. Recovery can continue an interrupted run under that same identity.
_Avoid_: Session or scenario as a synonym for run.

**Execution mode**: The full/retest classification of a run. A full run uses ordinary scenario selection, which may be filtered; a retest uses the retest selection and history workflow.
_Avoid_: Unqualified mode when transport or session mode could also be meant.

**Transport**: The route used to exchange testcase messages with the bot: WhatsApp or REST.
_Avoid_: Session mode as a synonym for transport.

**Session mode**: The conversation-context policy across scenarios: isolated or continuous. Turns within one scenario share context in either mode.

**Isolated session**: A session policy in which each scenario starts with independent bot context.

**Continuous session**: A session policy in which the selected scenarios share context across the run. Earlier scenarios can influence later responses.

### Results and recovery

**Technical status**: The outcome of sending and collecting a turn, such as Captured, Timeout, Send Error, or Chat Error. It describes execution rather than answer correctness.
_Avoid_: Passed or failed evaluation as a synonym for technical status.

**Semantic status**: The scenario's human-evaluation outcome or review stage, such as Passed, Failed, Ready for Re-test, or Pending Evaluation.
_Avoid_: Using captured as a synonym for passed.

**Evidence status**: The availability and publication state of screenshot evidence, independent of the technical and semantic statuses. Screenshot evidence is not applicable to REST execution.

**Captured scenario**: A scenario for which every turn has a Captured technical status. Its responses still require semantic evaluation.
_Avoid_: Passed scenario or completed run as a synonym.

**Retest**: Re-execution of selected scenarios with their prior results retained for comparison and new responses collected for evaluation.
_Avoid_: Automatic re-execution of every failed evaluation.

**Resume**: Continuation of an interrupted isolated run using its original selection and Run ID. An incomplete scenario restarts from its first turn.

**Continuous restart**: A new run that repeats an interrupted continuous run's entire original selection, in order, with fresh conversation context.
_Avoid_: Partial resume.
