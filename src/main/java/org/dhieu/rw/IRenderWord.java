package org.dhieu.rw;

import org.json.JSONObject;

import java.io.IOException;
import java.io.InputStream;

public interface IRenderWord {
    void render(JSONObject jsonObject, String urlTemplate, String output) throws IOException;

    void render(JSONObject jsonObject, InputStream input, String output) throws IOException;
}
