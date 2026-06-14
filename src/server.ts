import { buildApp } from "./app";

const port = Number(process.env.PORT ?? 3333);
const host = process.env.HOST ?? "0.0.0.0";

async function main(): Promise<void> {
  const app = buildApp({
    logger: process.env.API_LOGGER === "false" ? false : true
  });

  await app.listen({ port, host });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
