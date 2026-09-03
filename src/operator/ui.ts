import {
  cancel,
  confirm,
  intro,
  isCancel,
  log,
  note,
  outro,
  select,
  spinner,
  text,
  type SelectOptions,
} from "@clack/prompts";

export interface SelectOption<Value extends string> {
  value: Value;
  label: string;
  hint?: string;
  disabled?: boolean;
}

export interface SelectPrompt<Value extends string> {
  message: string;
  options: SelectOption<Value>[];
  initialValue?: Value;
}

export interface ConfirmPrompt {
  message: string;
  initialValue?: boolean;
  active?: string;
  inactive?: string;
}

export interface TextPrompt {
  message: string;
  placeholder?: string;
  initialValue?: string;
  defaultValue?: string;
  validate?: (value: string) => string | undefined;
}

export interface OperatorUi {
  intro(title: string): void;
  outro(message: string): void;
  cancel(message: string): void;
  select<Value extends string>(
    prompt: SelectPrompt<Value>,
  ): Promise<Value | undefined>;
  confirm(prompt: ConfirmPrompt): Promise<boolean | undefined>;
  text(prompt: TextPrompt): Promise<string | undefined>;
  note(message: string, title?: string): void;
  info(message: string): void;
  success(message: string): void;
  warn(message: string): void;
  error(message: string): void;
  task<Value>(
    message: string,
    operation: () => Promise<Value>,
    successMessage?: string,
  ): Promise<Value>;
}

export function createClackUi(): OperatorUi {
  return {
    intro,
    outro,
    cancel,
    async select<Value extends string>(prompt: SelectPrompt<Value>) {
      const result = await select<Value>({
        ...prompt,
        options: prompt.options as SelectOptions<Value>["options"],
        maxItems: 10,
      });
      return isCancel(result) ? undefined : result;
    },
    async confirm(prompt) {
      const result = await confirm(prompt);
      return isCancel(result) ? undefined : result;
    },
    async text(prompt) {
      const result = await text({
        ...prompt,
        validate: prompt.validate
          ? (value) => prompt.validate!(value ?? "")
          : undefined,
      });
      return isCancel(result) ? undefined : result;
    },
    note,
    info: log.info,
    success: log.success,
    warn: log.warn,
    error: log.error,
    async task<Value>(
      message: string,
      operation: () => Promise<Value>,
      successMessage?: string,
    ): Promise<Value> {
      const progress = spinner();
      progress.start(message);
      try {
        const result = await operation();
        progress.stop(successMessage ?? message);
        return result;
      } catch (error) {
        progress.error(`${message} failed`);
        throw error;
      }
    },
  };
}
