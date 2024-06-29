import { container } from "tsyringe";
import { EventEntry, LogEntry, MetricEntry, ILogger } from "../types";
import { AppInsightsService } from "./appInsights";
import { SeverityLevel } from "applicationinsights/out/Declarations/Contracts";
import { Env } from "../env";

export class AppInsightLogger implements ILogger {
  private appInsightClient: AppInsightsService | undefined = undefined;

  private env: Env;
  constructor() {
    this.env = container.resolve(Env);
    if (
      this.env.data &&
      this.env.data.APPLICATIONINSIGHTS_INSTRUMENTATION_KEY
    ) {
      this.appInsightClient = new AppInsightsService(
        this.env.data.APPLICATIONINSIGHTS_INSTRUMENTATION_KEY
      );
    }
  }

  public get isInitialized(): boolean {
    return this.appInsightClient !== undefined;
  }

  public getName(): string {
    return this.constructor.name;
  }

  private logWithSeverity(
    logEntry: LogEntry,
    severityLevel: SeverityLevel
  ): void {
    if (this.appInsightClient) {
      const msg = `[ApplicationInsights]: [${logEntry.module}]: [appVersion: ${this.env.data.APP_VERSION}] ${logEntry.message}`;
      this.appInsightClient.trackTrace(msg, severityLevel);
    }
  }

  public trace(logEntry: LogEntry): void {
    this.logWithSeverity(logEntry, SeverityLevel.Verbose);
  }

  public debug(logEntry: LogEntry): void {
    this.logWithSeverity(logEntry, SeverityLevel.Verbose);
  }

  public info(logEntry: LogEntry): void {
    this.logWithSeverity(logEntry, SeverityLevel.Information);
  }

  public warn(logEntry: LogEntry): void {
    this.logWithSeverity(logEntry, SeverityLevel.Warning);
  }

  public error(logEntry: LogEntry): void {
    this.logWithSeverity(logEntry, SeverityLevel.Error);
  }

  public trackEvent(eventEntry: EventEntry): void {
    if (this.appInsightClient) {
      this.appInsightClient.trackEvent(eventEntry.name, eventEntry.properties);
    }
  }

  public trackDurationMetric(metricEntry: MetricEntry): void {
    if (this.appInsightClient) {
      this.appInsightClient.trackDurationMetric(metricEntry);
    }
  }
}
