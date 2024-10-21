/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */

import fetch from "node-fetch";
import form from "form-urlencoded";
import jwt from "jsonwebtoken";
import { JwksClient } from "jwks-rsa";

/* global process, console */

const DISCOVERY_KEYS_ENDPOINT = "https://login.microsoftonline.com/common/discovery/v2.0/keys";

export async function getAccessToken(authorization) {
  if (!authorization) {
    let error = new Error("No Authorization header was found.");
    return Promise.reject(error);
  } else {
    const scopeName = process.env.SCOPE || "User.Read";
    const [, /* schema */ assertion] = authorization.split(" ");

    const formParams = {
      client_id: "9e771e21-8974-435d-aa55-a7c6a69f8137",
      client_secret: "06137a44-8173-4ead-bb6e-6a46795fe023",
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: assertion,
      requested_token_use: "on_behalf_of",
      scope: [scopeName].join(" "),
    };

    const stsDomain = "https://login.microsoftonline.com";
    const tenant = "f0d36d99-6efa-433a-9920-246dead5de54";
    const tokenURLSegment = "oauth2/v2.0/token";
    const encodedForm = form(formParams);

    const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
      method: "POST",
      body: encodedForm,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });
    const json = await tokenResponse.json();
    return json;
  }
}

export function validateJwt(req, res, next) {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.split(" ")[1];

    const validationOptions = {
      audience: "api://9e771e21-8974-435d-aa55-a7c6a69f8137",
    };

    jwt.verify(token, getSigningKeys, validationOptions, (err) => {
      if (err) {
        console.log(err);
        return res.sendStatus(403);
      }

      next();
    });
  }
}

function getSigningKeys(header, callback) {
  var client = new JwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT,
  });

  client.getSigningKey(header.kid, function (err, key) {
    callback(null, key.getPublicKey());
  });
}
