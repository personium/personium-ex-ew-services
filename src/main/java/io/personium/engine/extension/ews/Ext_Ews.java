/**
 * Personium
 * personium.io
 * Copyright 2017 FUJITSU LIMITED
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package io.personium.engine.extension.ews;

import io.personium.engine.extension.support.ExtensionErrorConstructor;
import io.personium.engine.extension.support.AbstractExtensionScriptableObject;
import io.personium.engine.extension.support.ExtensionLogger;
import org.mozilla.javascript.annotations.JSConstructor;
import org.mozilla.javascript.annotations.JSFunction;
import org.mozilla.javascript.NativeObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.property.complex.AttendeeCollection;
import microsoft.exchange.webservices.data.property.complex.Attendee;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;

import microsoft.exchange.webservices.data.core.enumeration.service.SendInvitationsOrCancellationsMode;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;

import java.io.IOException;
import java.net.Proxy;
import java.net.URI;
import java.util.Map;
import java.util.Date;
import java.util.TimeZone;
import java.util.ArrayList;
import java.util.Map.Entry;
import java.text.SimpleDateFormat;

import org.apache.commons.lang3.StringUtils;

/**
 * Engine-Extension EWS.
 */
@SuppressWarnings("serial")
public class Ext_Ews extends AbstractExtensionScriptableObject {

    static Logger log = LoggerFactory.getLogger(Ext_Ews.class);

    private ExchangeService service = null;


    /**
     * Public name to JavaScript.
     */
    @Override
    public String getClassName() {
        return "Ews";
    }


    /**
     * constructor.
     */
    @JSConstructor
    public Ext_Ews() {
        ExtensionLogger logger = new ExtensionLogger(this.getClass());
        setLogger(this.getClass(), logger);
    }


    /**
     *
     * @param emailAddress
     * @param password
     * @throws Exception
     */
    @JSFunction
    public void createService(String emailAddress, String password) throws Exception {

        if (service != null) {
            service = null;
        }
        service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        service.setTraceEnabled(true);
        ExchangeCredentials credentials = new WebCredentials(emailAddress, password);
        service.setCredentials(credentials);

    }


    /**
     *
     * @param emailAddress
     * @throws Exception
     */
    @JSFunction
    public String autodiscoverUrl(String emailAddress) throws Exception {

        service.autodiscoverUrl(emailAddress);
        return service.getUrl().toString();

    }


    /**
     *
     * @param serviceUrl
     * @throws Exception
     */
    @JSFunction
    public void setUrl(String serviceUrl) throws Exception {

        service.setUrl(new URI(serviceUrl));

    }


    /**
     *
     * @param reqStartDate(UTC: 2020-10-11T00:00:00.000Z)
     * @param reqEndDate(UTC: 2020-12-11T12:34:56.789Z)
     * @param maxNumber
     * @throws Exception
     */
    @JSFunction
    public ArrayList<NativeObject> findVEvents(String reqStartDate, String reqEndDate, String maxNumber) throws Exception {

      ArrayList<NativeObject> results = new ArrayList<NativeObject>();
      SimpleDateFormat utcFormat = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
      utcFormat.setTimeZone(TimeZone.getTimeZone("UTC"));

      try {
          Date startDate = utcFormat.parse(reqStartDate);
          Date endDate = utcFormat.parse(reqEndDate);
          int maxReturned = Integer.parseInt(maxNumber);
          CalendarFolder cf = CalendarFolder.bind(service, WellKnownFolderName.Calendar);
          FindItemsResults<Appointment> findResults = cf.findAppointments(new CalendarView(startDate, endDate, maxReturned));

          // Set NativeObject.
          results = new ArrayList<NativeObject>();
          NativeObject result = null;

          for (Appointment appt : findResults.getItems()) {
              appt.load(PropertySet.FirstClassProperties);
              result = new NativeObject();

              result.put("Uid", result, appt.getId().getUniqueId());
              result.put("ICalUid", result, appt.getICalUid());
              result.put("Subject", result, appt.getSubject());
              result.put("Start", result, utcFormat.format(appt.getStart()).toString());
              result.put("End", result, utcFormat.format(appt.getEnd()).toString());
              result.put("Body", result, appt.getBody().toString());
              result.put("Location", result, appt.getLocation());
              result.put("Organizer", result, appt.getOrganizer().getAddress());
              ArrayList<String> attendees = new ArrayList<String>();
              for (Attendee attendee : appt.getRequiredAttendees().getItems()) {
                  attendees.add(attendee.getAddress());
              }
              result.put("Attendees", result, attendees.toString());
              result.put("Updated", result, utcFormat.format(appt.getICalDateTimeStamp()).toString());

              results.add(result);

          }


        } catch (Exception e) {
            String message = "An error occurred.";
            this.getLogger().warn(message, e);
            String errorMessage = String.format("%s Cause: [%s]",
                    message, e.getClass().getName() + ": " + e.getMessage());
            throw ExtensionErrorConstructor.construct(errorMessage);
        }
        return results;
    }


    /**
     *
     * @param vEvent { key = (srcId, dtstart, dtend, summary, location, description, attendees)}
     * @throws Exception
     */
    @JSFunction
    public NativeObject updateVEvent(NativeObject vEvent) throws Exception {

        SimpleDateFormat utcFormat = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
        utcFormat.setTimeZone(TimeZone.getTimeZone("UTC"));

        Appointment appointment = null;
        NativeObject result = new NativeObject();
        String Uid = null;

        try {

            for (Entry<Object, Object> ve : vEvent.entrySet()) {
                if (ve.getKey().toString().equals("srcId")) {
                    Uid = ve.getValue().toString();
                    appointment = Appointment.bind(service, new ItemId(Uid));
                }
            }

            appointment.setSubject("");
            appointment.setLocation("");
            appointment.setBody(MessageBody.getMessageBodyFromText(""));
            appointment.getRequiredAttendees().clear();

            for (Entry<Object, Object> ve : vEvent.entrySet()) {
                switch (ve.getKey().toString()) {
                case "dtstart":
                    Date startDate = utcFormat.parse(ve.getValue().toString());
                    appointment.setStart(startDate);
                    break;
                case "dtend":
                    Date endDate = utcFormat.parse(ve.getValue().toString());
                    appointment.setEnd(endDate);
                    break;
                case "summary":
                    appointment.setSubject(ve.getValue().toString());
                    break;
                case "location":
                    appointment.setLocation(ve.getValue().toString());
                    break;
                case "description":
                    appointment.setBody(MessageBody.getMessageBodyFromText(ve.getValue().toString()));
                    break;
                case "attendees":
                    this.getLogger().info(ve.getValue().toString());
                    String[] atten = StringUtils.split(StringUtils.deleteWhitespace(StringUtils.strip(ve.getValue().toString(), "[]")), ",");
                    for (int i=0; i < atten.length; i++) {
                        appointment.getRequiredAttendees().add(atten[i]);
                    }
                    break;
                    //default:
                   }
               }

               //appointment.update(ConflictResolutionMode.AlwaysOverwrite, SendInvitationsOrCancellationsMode.SendToNone);
               appointment.update(ConflictResolutionMode.AutoResolve);

               Appointment appt = Appointment.bind(service, new ItemId(Uid));

               result.put("Uid", result, appt.getId().getUniqueId());
               result.put("ICalUid", result, appt.getICalUid());
               result.put("Subject", result, appt.getSubject());
               result.put("Start", result, utcFormat.format(appt.getStart()).toString());
               result.put("End", result, utcFormat.format(appt.getEnd()).toString());
               result.put("Body", result, appt.getBody().toString());
               result.put("Location", result, appt.getLocation());
               result.put("Organizer", result, appt.getOrganizer().getAddress());
               ArrayList<String> attendees = new ArrayList<String>();
               for (Attendee attendee : appt.getRequiredAttendees().getItems()) {
                   attendees.add(attendee.getAddress());
               }
               result.put("Attendees", result, attendees.toString());
               result.put("Updated", result, utcFormat.format(appt.getICalDateTimeStamp()).toString());


           } catch (Exception e) {
               String message = "An error occurred.";
               this.getLogger().warn(message, e);
               String errorMessage = String.format("%s Cause: [%s]",
                       message, e.getClass().getName() + ": " + e.getMessage());
               throw ExtensionErrorConstructor.construct(errorMessage);
           }

           return result;
       }


       /**
        *
        * @param vEvent { key = (srcId, dtstart, dtend, summary, location, description, attendees)}
        * @throws Exception
        */
       @JSFunction
       public String deleteVEvent(NativeObject vEvent) throws Exception {

           Appointment appointment = null;
           String Uid = null;

           try {

               for (Entry<Object, Object> ve : vEvent.entrySet()) {
                   if (ve.getKey().toString().equals("srcId")) {
                       Uid = ve.getValue().toString();
                       appointment = Appointment.bind(service, new ItemId(Uid));
                   }
               }

               appointment.delete(DeleteMode.MoveToDeletedItems);

           } catch (Exception e) {
               String message = "An error occurred.";
               this.getLogger().warn(message, e);
               String errorMessage = String.format("%s Cause: [%s]",
                       message, e.getClass().getName() + ": " + e.getMessage());
               throw ExtensionErrorConstructor.construct(errorMessage);
           }

           return "OK";
       }


}
