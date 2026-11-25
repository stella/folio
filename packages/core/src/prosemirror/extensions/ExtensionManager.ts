/**
 * Extension Manager
 *
 * Two-phase initialization:
 * 1. buildSchema() — collects NodeSpecs/MarkSpecs from extensions → new Schema
 * 2. initializeRuntime() — calls onSchemaReady() on each extension, collects plugins/commands/keymaps
 */

import { panic } from "better-result";
import { Schema } from "prosemirror-model";
import type { MarkSpec, NodeSpec } from "prosemirror-model";
import type { Plugin as PMPlugin } from "prosemirror-state";

import { createExtensionRuntime } from "../../internal/extensionRuntime";
import type {
  AnyExtension,
  CommandFactory,
  CommandMap,
  ExtensionCommandMap,
  FolioCommandArguments,
  FolioCommandName,
} from "./types";

export class ExtensionManager {
  private readonly extensions: AnyExtension[];
  private schema: Schema | null = null;
  private plugins: PMPlugin[] = [];
  private commands: ExtensionCommandMap = {};

  constructor(extensions: AnyExtension[]) {
    // Sort by priority (lower number = higher priority)
    this.extensions = [...extensions].toSorted(
      (a, b) => a.config.priority - b.config.priority,
    );
  }

  /**
   * Phase 1: Build schema from node/mark extensions
   */
  buildSchema(): void {
    const nodes: Record<string, NodeSpec> = {};
    const marks: Record<string, MarkSpec> = {};

    for (const ext of this.extensions) {
      if (ext.type === "node") {
        nodes[ext.config.schemaNodeName] = ext.config.nodeSpec;
      } else if (ext.type === "mark") {
        marks[ext.config.schemaMarkName] = ext.config.markSpec;
      }
    }

    this.schema = new Schema({ nodes, marks });
  }

  /**
   * Phase 2: Initialize runtime (plugins, commands, keymaps)
   * Must be called after buildSchema()
   */
  initializeRuntime(): void {
    if (!this.schema) {
      panic(
        "ExtensionManager: buildSchema() must be called before initializeRuntime()",
      );
    }

    const runtime = createExtensionRuntime(this.extensions, this.schema);
    this.plugins = runtime.plugins;
    this.commands = runtime.commandRegistry;
  }

  /**
   * Get the built schema
   */
  getSchema(): Schema {
    if (!this.schema) {
      panic("ExtensionManager: buildSchema() must be called first");
    }
    return this.schema;
  }

  /**
   * Get all plugins (raw + keymap merged)
   */
  getPlugins(): PMPlugin[] {
    return this.plugins;
  }

  /**
   * Get the flat command registry
   */
  getCommands(): CommandMap {
    return this.commands;
  }

  /**
   * Get a specific command by name
   */
  getCommand<Name extends FolioCommandName>(
    name: Name,
  ): CommandFactory<FolioCommandArguments[Name]> | undefined;
  getCommand(name: string): CommandFactory | undefined;
  getCommand(name: string): CommandFactory | undefined {
    return this.commands[name];
  }

  requireCommand<Name extends FolioCommandName>(
    name: Name,
  ): CommandFactory<FolioCommandArguments[Name]>;
  requireCommand(name: string): CommandFactory;
  requireCommand(name: string): CommandFactory {
    const command = this.getCommand(name);
    if (!command) {
      panic(`ExtensionManager: command "${name}" is not registered`);
    }
    return command;
  }

  /**
   * Lifecycle: destroy
   */
  destroy(): void {
    this.plugins = [];
    this.commands = {};
  }
}
