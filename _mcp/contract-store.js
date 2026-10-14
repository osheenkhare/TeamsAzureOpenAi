const fs = require("node:fs");
const path = require("node:path");
const { z } = require("zod");

const contractsPath = path.join(__dirname, "contracts.json");
const textSchema = z.string().trim().min(1);
const contractFields = {
  title: textSchema,
  description: textSchema,
  status: z.enum(["open", "in-progress", "approved", "blocked"]),
  paymentTerm: z.enum(["net-10", "net-30", "net-60", "net-80"]),
  createdByUser: textSchema,
  requireApproval: z.boolean(),
  approvalStatus: z.array(z.object({
    user: textSchema,
    approvalStatus: z.enum(["pending", "approved", "rejected"])
  }).strict())
};
const contractSchema = z.object({ id: textSchema, ...contractFields }).strict();

function readContracts() {
  const contracts = z.array(contractSchema).parse(JSON.parse(fs.readFileSync(contractsPath, "utf8")));
  if (new Set(contracts.map((contract) => contract.id)).size !== contracts.length) {
    throw new Error("Contract IDs must be unique.");
  }
  return contracts;
}

function writeContracts(contracts) {
  const temporaryPath = `${contractsPath}.${process.pid}.tmp`;
  fs.writeFileSync(temporaryPath, `${JSON.stringify(contracts, null, 2)}\n`);
  fs.renameSync(temporaryPath, contractsPath);
}

module.exports = { textSchema, contractFields, contractSchema, readContracts, writeContracts };