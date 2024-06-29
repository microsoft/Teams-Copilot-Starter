/**
 * SampleDocument model
 */
export interface SampleDocument {
  id?: string;
  title?: string | null;
  content?: string | null;
  contentVector?: number[] | null;
  filepath?: string | null;
  url?: string | null;
  last_updated?: string | null;
}
