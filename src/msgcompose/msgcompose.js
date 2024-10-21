/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { getUserProfile } from "src/helpers/sso-helper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getProfileButton").onclick = run;
  }
});

export async function run() {
  await getUserProfile(showMsgComposeTaskPane);
}

async function showMsgComposeTaskPane(societesData) {
  console.log(societesData);
}
