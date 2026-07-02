#!/usr/bin/env node
import { runCli } from '../dist/migrate.js'
process.exit(runCli(process.argv.slice(2)))
