import { Client } from "@microsoft/microsoft-graph-client";

// Shared shape of a single page from any Graph collection endpoint.
// `@odata.nextLink`, when present, is the absolute URL of the next page —
// already carrying the original query/select/top and the skip token.
export type GraphPage<T> = {
  value: T[];
  "@odata.nextLink"?: string;
};

/**
 * Fetches a Graph collection endpoint honoring the `unroll` option shared by
 * every paginated method on GraphApiService: `false`/omitted returns just
 * that page (plus `@odata.nextLink` if there's more); `true` follows every
 * subsequent page and returns everything aggregated, with no
 * `@odata.nextLink` left on the result.
 */
export async function fetchGraphPage<T>(
  client: Client,
  url: string,
  unroll?: boolean,
): Promise<GraphPage<T>> {
  const firstPage: GraphPage<T> = await client.api(url).get();
  if (!unroll) return firstPage;

  const value = [...firstPage.value];
  let nextLink = firstPage["@odata.nextLink"];
  while (nextLink) {
    const page: GraphPage<T> = await client.api(nextLink).get();
    value.push(...page.value);
    nextLink = page["@odata.nextLink"];
  }
  return { value };
}

/**
 * Fetches the next page of any paginated collection call made without
 * `unroll: true` — pass the `@odata.nextLink` from the previous response
 * verbatim, regardless of which method produced it (sites, users, ...).
 */
export function fetchGraphNextPage<T>(
  client: Client,
  nextLink: string,
): Promise<GraphPage<T>> {
  return client.api(nextLink).get();
}
